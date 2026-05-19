from __future__ import annotations

import asyncio
import math
import os
import secrets
import time
import uuid
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, AsyncIterator, Dict, List, Optional, Set, Tuple

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pydantic import BaseModel, Field

from .tool_registry import get_tool_specs


class McpExecuteRequest(BaseModel):
  tool_name: str
  tool_input: Dict[str, Any] = Field(default_factory=dict)


class McpToolResultRequest(BaseModel):
  request_id: str
  nonce: str
  delivery_id: Optional[str] = None
  ack: bool = False
  result: Optional[Dict[str, Any]] = None
  error: Optional[Dict[str, Any]] = None


_TOOL_TIMEOUT = int(os.getenv("EXCEL_MCP_TOOL_TIMEOUT", "180"))
RELAY_RECONNECT_GRACE_S = int(os.getenv("RELAY_RECONNECT_GRACE_S", str(_TOOL_TIMEOUT + 30)))


@dataclass
class ClientState:
  gateway_session_id: str
  user_id: str
  workbook_session: str
  client_id: str
  workbook_name: str
  queue: "asyncio.Queue[Dict[str, Any]]"
  connected_at: float
  detach_grace_deadline: Optional[float] = None
  detached_at: Optional[float] = None

  @property
  def detached(self) -> bool:
    return self.detach_grace_deadline is not None


class McpRelay:
  ACK_REPLAY_GRACE_SECONDS = 5
  EXPIRED_TTL_SECONDS = 300

  def __init__(self) -> None:
    self._lock = asyncio.Lock()
    self._inflight: Dict[str, Dict[str, Any]] = {}
    self._clients: Dict[str, ClientState] = {}
    self._active_workbook: Dict[str, str] = {}
    self._session_tools: Dict[str, List[str]] = {}
    self._background_tasks: Set["asyncio.Task[Any]"] = set()
    self.reconnect_grace_seconds = RELAY_RECONNECT_GRACE_S

  async def execute(
    self,
    tool_name: str,
    tool_input: Dict[str, Any],
    timeout: int,
    target_session: Optional[str] = None,
    *,
    gateway_session_id: Optional[str] = None,
    user_id: Optional[str] = None,
  ) -> Dict[str, Any]:
    request_id = uuid.uuid4().hex
    nonce = os.urandom(12).hex()
    result_future: "asyncio.Future[Dict[str, Any]]" = asyncio.get_running_loop().create_future()
    tool_payload = dict(tool_input)
    requested_session = target_session or tool_payload.pop("_workbook", None)

    async with self._lock:
      self._prune_expired_locked()
      if requested_session is not None and not isinstance(requested_session, str):
        raise RuntimeError("unknown_session")
      client = self._resolve_client_locked(gateway_session_id, requested_session)
      if client is None:
        if gateway_session_id is None and requested_session is None:
          raise RuntimeError("no_active_session")
        raise RuntimeError("unknown_session" if requested_session is not None else "session_disconnected")
      if user_id is not None and client.user_id != user_id:
        raise RuntimeError("session_disconnected")
      if self._client_detached_past_grace_locked(client):
        raise RuntimeError("session_disconnected")

      entry = {
        "request_id": request_id,
        "nonce": nonce,
        "tool_name": tool_name,
        "tool_input": tool_payload,
        "state": "queued",
        "created_at": time.time(),
        "delivered_at": None,
        "acked_at": None,
        "timeout": timeout,
        "session_id": client.gateway_session_id,
        "user_id": client.user_id,
        "workbook_session": client.workbook_session,
        "target_session": client.gateway_session_id,
        "delivery_client_id": None,
        "future": result_future,
      }
      self._inflight[request_id] = entry

    delivered = await self._deliver(request_id, replay=False)
    if not delivered:
      async with self._lock:
        self._inflight.pop(request_id, None)
      raise RuntimeError("session_disconnected")

    try:
      result_payload = await asyncio.wait_for(asyncio.shield(result_future), timeout=timeout)
    except asyncio.TimeoutError:
      async with self._lock:
        existing = self._inflight.get(request_id)
        if existing:
          existing["state"] = "expired"
          existing["expired_at"] = time.time()
      raise

    async with self._lock:
      self._inflight.pop(request_id, None)

    return {
      "request_id": request_id,
      "result": result_payload.get("result"),
      "error": result_payload.get("error"),
    }

  async def list_workbooks(
    self,
    gateway_session_id: Optional[str] = None,
    user_id: Optional[str] = None,
  ) -> List[Dict[str, Any]]:
    async with self._lock:
      self._prune_expired_locked()
      clients = self._clients.values()
      if gateway_session_id is not None:
        clients = [
          client
          for client in clients
          if client.gateway_session_id == gateway_session_id
          and (user_id is None or client.user_id == user_id)
        ]
      workbooks = [
        {
          "name": client.workbook_name,
          "session": client.workbook_session,
          "gateway_session_id": client.gateway_session_id,
          "active": self._active_workbook.get(client.gateway_session_id) == client.workbook_session,
          "connected_at": client.connected_at,
          "detached": client.detached,
          "detach_grace_deadline": client.detach_grace_deadline,
        }
        for client in sorted(clients, key=lambda item: item.connected_at)
      ]
    return workbooks

  async def switch_active_workbook(
    self,
    session: str,
    *,
    gateway_session_id: Optional[str] = None,
    user_id: Optional[str] = None,
  ) -> Dict[str, Any]:
    async with self._lock:
      self._prune_expired_locked()
      client = self._resolve_client_locked(gateway_session_id, session)
      if client is None:
        raise RuntimeError("unknown_session")
      if user_id is not None and client.user_id != user_id:
        raise RuntimeError("unknown_session")
      self._active_workbook[client.gateway_session_id] = client.workbook_session

    await client.queue.put({"type": "active_changed", "new_active": client.workbook_session})
    return {"status": "ok", "active": client.workbook_session}

  async def register_client(
    self,
    session: Optional[str],
    workbook_name: Optional[str],
    *,
    gateway_session_id: Optional[str] = None,
    user_id: Optional[str] = None,
    workbook_session: Optional[str] = None,
    tool_names: Optional[Set[str]] = None,
  ) -> ClientState:
    session_token = str(gateway_session_id or session or "").strip()
    if not session_token:
      raise ValueError("gateway_session_id or session is required")
    if session_token.startswith("legacy:"):
      raise ValueError("legacy session tokens are no longer supported")
    effective_user_id = str(user_id or "").strip()
    if not effective_user_id:
      raise ValueError("user_id is required")

    queue: "asyncio.Queue[Dict[str, Any]]" = asyncio.Queue()
    client_id = uuid.uuid4().hex
    now = time.time()
    effective_session = session_token
    effective_workbook_session = workbook_session or session or effective_session
    display_name = workbook_name or "unknown"
    client = ClientState(
      gateway_session_id=effective_session,
      user_id=effective_user_id,
      workbook_session=effective_workbook_session,
      client_id=client_id,
      workbook_name=display_name,
      queue=queue,
      connected_at=now,
    )
    replaced_queue: Optional["asyncio.Queue[Dict[str, Any]]"] = None
    immediate_replays: List[Tuple["asyncio.Queue[Dict[str, Any]]", Dict[str, Any]]] = []
    delayed_replays: List[Tuple[str, str, float]] = []

    async with self._lock:
      self._prune_expired_locked()
      replaced = self._clients.get(effective_session)
      if replaced is not None:
        if user_id is not None and replaced.user_id != user_id:
          raise ValueError("Session already registered for a different user")
        if not replaced.detached:
          replaced_queue = replaced.queue
      self._clients[effective_session] = client
      self._active_workbook[effective_session] = effective_workbook_session
      self._session_tools[effective_session] = sorted(tool_names or [])

      for request_id, item in self._inflight.items():
        if item.get("state") != "delivered":
          continue
        if item.get("session_id") != effective_session:
          continue
        expected_prior_delivery = item.get("delivery_client_id")
        if not isinstance(expected_prior_delivery, str):
          continue
        delivered_at = item.get("delivered_at") or 0
        age = now - float(delivered_at)
        if age >= self.ACK_REPLAY_GRACE_SECONDS:
          if item.get("delivery_client_id") != expected_prior_delivery:
            continue
          payload = self._mark_delivered_locked(
            request_id,
            item,
            client_id,
            replay=True,
          )
          immediate_replays.append((queue, payload))
        else:
          delayed_replays.append((request_id, expected_prior_delivery, self.ACK_REPLAY_GRACE_SECONDS - age))

    if replaced_queue is not None:
      await replaced_queue.put(
        {
          "type": "replaced",
          "reason": "Same-session reconnect",
        }
      )

    for replay_queue, payload in immediate_replays:
      await replay_queue.put(payload)

    for request_id, expected_prior_delivery, delay in delayed_replays:
      self._schedule_delayed_replay(effective_session, request_id, expected_prior_delivery, delay)

    return client

  async def unregister_client(self, session: str, client_id: str, *, immediate: bool = False) -> None:
    async with self._lock:
      client = self._clients.get(session)
      if client is None or client.client_id != client_id:
        return
      if immediate:
        self._remove_client_locked(session)
        return
      now = time.time()
      client.detached_at = now
      client.detach_grace_deadline = now + self.reconnect_grace_seconds
      self._schedule_detached_sweep(session, client_id, self.reconnect_grace_seconds)

  async def ack(
    self,
    request_id: str,
    nonce: str,
    delivery_id: Optional[str],
    *,
    gateway_session_id: Optional[str] = None,
    user_id: Optional[str] = None,
  ) -> str:
    async with self._lock:
      self._prune_expired_locked()
      entry = self._inflight.get(request_id)
      if not entry:
        return "not_found"
      ownership_status = self._validate_owner_locked(entry, gateway_session_id, user_id)
      if ownership_status != "ok":
        return ownership_status
      if entry.get("nonce") != nonce:
        return "nonce_mismatch"
      delivery_status = self._validate_delivery_id_locked(entry, delivery_id)
      if delivery_status != "ok":
        return delivery_status
      if entry.get("state") == "orphaned":
        return "expired"
      if entry.get("state") == "expired":
        self._inflight.pop(request_id, None)
        return "expired"
      if entry.get("state") == "completed":
        return "completed"
      if entry.get("state") == "acked":
        return "ok"
      entry["state"] = "acked"
      entry["acked_at"] = time.time()
      return "ok"

  async def complete(
    self,
    request_id: str,
    nonce: str,
    delivery_id: Optional[str],
    result: Optional[Dict[str, Any]],
    error: Optional[Dict[str, Any]],
    *,
    gateway_session_id: Optional[str] = None,
    user_id: Optional[str] = None,
  ) -> str:
    result_future: Optional["asyncio.Future[Dict[str, Any]]"] = None

    async with self._lock:
      self._prune_expired_locked()
      entry = self._inflight.get(request_id)
      if not entry:
        return "not_found"
      ownership_status = self._validate_owner_locked(entry, gateway_session_id, user_id)
      if ownership_status != "ok":
        return ownership_status
      if entry.get("nonce") != nonce:
        return "nonce_mismatch"
      delivery_status = self._validate_delivery_id_locked(entry, delivery_id)
      if delivery_status != "ok":
        return delivery_status
      if entry.get("state") == "orphaned":
        return "expired"
      if entry.get("state") == "expired":
        self._inflight.pop(request_id, None)
        return "expired"
      if entry.get("state") == "completed":
        return "completed"
      entry["state"] = "completed"
      entry["completed_at"] = time.time()
      result_future = entry.get("future")

    if result_future is not None and not result_future.done():
      result_future.set_result({"result": result, "error": error})

    return "ok"

  async def status(
    self,
    gateway_session_id: str,
    user_id: Optional[str] = None,
  ) -> Dict[str, Any]:
    async with self._lock:
      self._prune_expired_locked()
      client = self._clients.get(gateway_session_id)
      if client is not None and user_id is not None and client.user_id != user_id:
        client = None
      inflight_count = sum(
        1
        for item in self._inflight.values()
        if item.get("session_id") == gateway_session_id
        and (user_id is None or item.get("user_id") == user_id)
        and item.get("state") != "expired"
      )
      return {
        "session_id": gateway_session_id,
        "connected": bool(client and not client.detached),
        "detached": bool(client and client.detached),
        "workbook_session": client.workbook_session if client else None,
        "workbook_name": client.workbook_name if client else None,
        "active_workbook": self._active_workbook.get(gateway_session_id),
        "tools": self._session_tools.get(gateway_session_id, []),
        "inflight": inflight_count,
        "detach_grace_deadline": client.detach_grace_deadline if client else None,
      }

  async def _deliver(self, request_id: str, replay: bool) -> bool:
    async with self._lock:
      entry = self._inflight.get(request_id)
      if not entry:
        return False
      target_session = entry.get("session_id")
      if not isinstance(target_session, str):
        return False
      client = self._clients.get(target_session)
      if client is None or self._client_detached_past_grace_locked(client):
        return False
      payload = self._mark_delivered_locked(
        request_id,
        entry,
        client.client_id,
        replay=replay,
      )
      queue = client.queue

    await queue.put(payload)
    return True

  def _schedule_delayed_replay(
    self,
    session: str,
    request_id: str,
    expected_prior_delivery: str,
    delay: float,
  ) -> None:
    task: Optional["asyncio.Task[Any]"] = None

    async def _runner() -> None:
      try:
        await asyncio.sleep(max(delay, 0))
        async with self._lock:
          entry = self._inflight.get(request_id)
          if (
            not entry
            or entry.get("state") != "delivered"
            or entry.get("session_id") != session
            or entry.get("delivery_client_id") != expected_prior_delivery
          ):
            return
          client = self._clients.get(session)
          if client is None or self._client_detached_past_grace_locked(client):
            return
          payload = self._mark_delivered_locked(
            request_id,
            entry,
            client.client_id,
            replay=True,
          )
          queue = client.queue
        await queue.put(payload)
      finally:
        if task is not None:
          self._background_tasks.discard(task)

    task = asyncio.create_task(_runner())
    self._background_tasks.add(task)

  def _schedule_detached_sweep(self, session: str, client_id: str, delay: float) -> None:
    task: Optional["asyncio.Task[Any]"] = None

    async def _runner() -> None:
      try:
        await asyncio.sleep(max(delay, 0))
        async with self._lock:
          client = self._clients.get(session)
          if client is None or client.client_id != client_id:
            return
          if not self._client_detached_past_grace_locked(client):
            return
          self._remove_client_locked(session, orphan_inflight=True)
      finally:
        if task is not None:
          self._background_tasks.discard(task)

    task = asyncio.create_task(_runner())
    self._background_tasks.add(task)

  def _mark_delivered_locked(
    self,
    request_id: str,
    entry: Dict[str, Any],
    delivery_id: str,
    replay: bool,
  ) -> Dict[str, Any]:
    entry["state"] = "delivered"
    entry["delivered_at"] = time.time()
    entry["delivery_client_id"] = delivery_id
    return {
      "type": "mcp_tool_request",
      "request_id": request_id,
      "nonce": entry.get("nonce"),
      "delivery_id": delivery_id,
      "tool_name": entry.get("tool_name"),
      "tool_input": entry.get("tool_input"),
      "replay": replay,
    }

  def _validate_delivery_id_locked(
    self,
    entry: Dict[str, Any],
    delivery_id: Optional[str],
  ) -> str:
    if delivery_id is None:
      return "delivery_id_required"
    return "ok" if entry.get("delivery_client_id") == delivery_id else "stale_delivery"

  def _validate_owner_locked(
    self,
    entry: Dict[str, Any],
    gateway_session_id: Optional[str],
    user_id: Optional[str],
  ) -> str:
    if gateway_session_id is not None and entry.get("session_id") != gateway_session_id:
      return "session_mismatch"
    if user_id is not None and entry.get("user_id") != user_id:
      return "user_mismatch"
    return "ok"

  def _resolve_client_locked(
    self,
    gateway_session_id: Optional[str],
    requested_session: Optional[str],
  ) -> Optional[ClientState]:
    if gateway_session_id is not None:
      client = self._clients.get(gateway_session_id)
      if client is None:
        return None
      if requested_session and requested_session not in {
        client.gateway_session_id,
        client.workbook_session,
      }:
        return None
      return client

    if requested_session:
      client = self._clients.get(requested_session)
      if client is not None:
        return client
      for candidate in self._clients.values():
        if candidate.workbook_session == requested_session:
          return candidate
      return None

    active_gateway_session = next(iter(self._active_workbook.keys()), None)
    if active_gateway_session is None:
      return None
    return self._clients.get(active_gateway_session)

  def _client_detached_past_grace_locked(self, client: ClientState) -> bool:
    deadline = client.detach_grace_deadline
    return deadline is not None and time.time() >= deadline

  def _remove_client_locked(self, session: str, *, orphan_inflight: bool = False) -> None:
    self._clients.pop(session, None)
    self._active_workbook.pop(session, None)
    self._session_tools.pop(session, None)
    if orphan_inflight:
      for item in self._inflight.values():
        if item.get("session_id") == session and item.get("state") in {"queued", "delivered", "acked"}:
          item["state"] = "orphaned"
          item["orphaned_at"] = time.time()

  def _prune_expired_locked(self) -> None:
    now = time.time()
    to_delete: List[str] = []
    for request_id, item in self._inflight.items():
      if item.get("state") != "expired":
        continue
      expired_at = float(item.get("expired_at") or now)
      if now - expired_at >= self.EXPIRED_TTL_SECONDS:
        to_delete.append(request_id)
    for request_id in to_delete:
      self._inflight.pop(request_id, None)

    stale_sessions = [
      session
      for session, client in self._clients.items()
      if client.detach_grace_deadline is not None and now >= client.detach_grace_deadline
    ]
    for session in stale_sessions:
      self._remove_client_locked(session, orphan_inflight=True)


class ChannelType(Enum):
  EXCEL = "excel"
  LOCAL = "local"
  MCP_EXTERNAL = "mcp_ext"


@dataclass
class Channel:
  channel_id: str
  channel_type: ChannelType
  tool_names: Set[str]
  connected_at: float = field(default_factory=time.time)
  metadata: Dict[str, Any] = field(default_factory=dict)


class ChannelRegistry:
  def __init__(self) -> None:
    self._lock = asyncio.Lock()
    self._channels: Dict[str, Channel] = {}
    self._tool_to_channel: Dict[str, str] = {}

  async def register(self, channel: Channel) -> None:
    async with self._lock:
      self._channels[channel.channel_id] = channel
      self._rebuild_tool_index_locked()

  async def unregister(self, channel_id: str) -> None:
    async with self._lock:
      self._channels.pop(channel_id, None)
      self._rebuild_tool_index_locked()

  def get_channel_for_tool(self, tool_name: str) -> Optional[Channel]:
    channel_id = self._tool_to_channel.get(tool_name)
    if channel_id is None:
      return None
    return self._channels.get(channel_id)

  def get_available_tool_names(self) -> Set[str]:
    return set(self._tool_to_channel.keys())

  def get_active_channels(self) -> List[Channel]:
    return list(self._channels.values())

  def is_channel_type_connected(self, channel_type: ChannelType) -> bool:
    return any(channel.channel_type == channel_type for channel in self._channels.values())

  def _rebuild_tool_index_locked(self) -> None:
    tool_to_channel: Dict[str, str] = {}
    for channel in self._channels.values():
      for tool_name in channel.tool_names:
        tool_to_channel[tool_name] = channel.channel_id
    self._tool_to_channel = tool_to_channel


class RelayState:
  def __init__(self, tool_names: Set[str]) -> None:
    self.relay = McpRelay()
    self.channel_registry = ChannelRegistry()
    self.tool_names = set(tool_names)
    self.cleanup_tasks: Set["asyncio.Task[None]"] = set()


def _sanitize_for_json(obj: Any) -> Any:
  """Recursively replace NaN/Infinity float values with None."""
  if isinstance(obj, float) and not math.isfinite(obj):
    return None
  if isinstance(obj, dict):
    return {key: _sanitize_for_json(value) for key, value in obj.items()}
  if isinstance(obj, (list, tuple)):
    return [_sanitize_for_json(value) for value in obj]
  if isinstance(obj, (set, frozenset)):
    return [_sanitize_for_json(value) for value in obj]
  return obj


def json_dumps(payload: Dict[str, Any]) -> str:
  """Serialize JSON using FastAPI's response encoder for SSE safety."""
  sanitized = _sanitize_for_json(payload)
  return JSONResponse(content=sanitized).body.decode("utf-8")


def _is_valid_mcp_secret(candidate: str, secret: str) -> bool:
  if os.getenv("ENVIRONMENT", "").strip().lower() != "development":
    return False
  if not secret:
    return False
  return secrets.compare_digest(candidate or "", secret)


def create_relay_app(
  secret_env_var: str = "EXCEL_MCP_SECRET",
  cors_origins: Optional[List[str]] = None,
  tool_names: Optional[Set[str]] = None,
) -> FastAPI:
  if cors_origins is None:
    cors_raw = os.getenv("CHAT_CORS_ORIGINS", "https://localhost:3002,http://localhost:3002")
    cors_origins = [origin.strip() for origin in cors_raw.split(",") if origin.strip()]

  if tool_names is None:
    tool_names = {spec["name"] for spec in get_tool_specs()}

  state = RelayState(tool_names=set(tool_names))
  app = FastAPI()
  app.state.mcp_relay_state = state

  app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Authorization", "Content-Type", "X-MCP-Secret"],
  )

  def _secret() -> str:
    return os.getenv(secret_env_var, "").strip()

  def _relay_error_response(code: str) -> JSONResponse:
    if code == "no_active_session":
      return JSONResponse(
        {"error": "No active Excel taskpane connection", "code": code},
        status_code=503,
      )
    if code == "unknown_session":
      return JSONResponse(
        {"error": "Unknown workbook session", "code": code},
        status_code=404,
      )
    if code == "session_disconnected":
      return JSONResponse(
        {"error": "Target workbook session disconnected", "code": code},
        status_code=503,
      )
    return JSONResponse({"error": f"MCP relay failed: {code}"}, status_code=500)

  @app.get("/api/health")
  async def health() -> JSONResponse:
    return JSONResponse({"status": "ok"})

  @app.get("/api/mcp/channel-status")
  async def mcp_channel_status(request: Request) -> JSONResponse:
    secret = request.headers.get("X-MCP-Secret", "")
    if not _is_valid_mcp_secret(secret, _secret()):
      return JSONResponse({"error": "Unauthorized"}, status_code=401)
    channels = state.channel_registry.get_active_channels()
    return JSONResponse(
      {
        "excel_connected": state.channel_registry.is_channel_type_connected(ChannelType.EXCEL),
        "channels": [
          {"type": channel.channel_type.value, "id": channel.channel_id}
          for channel in channels
        ],
      }
    )

  @app.post("/api/mcp/execute")
  async def mcp_execute(request: Request, payload: McpExecuteRequest) -> JSONResponse:
    secret = request.headers.get("X-MCP-Secret", "")
    if not _is_valid_mcp_secret(secret, _secret()):
      return JSONResponse({"error": "Unauthorized MCP secret"}, status_code=401)

    timeout = int(os.getenv("EXCEL_MCP_TOOL_TIMEOUT", "60"))
    if payload.tool_name == "list_workbooks":
      return JSONResponse({"workbooks": await state.relay.list_workbooks()})
    if payload.tool_name == "switch_active_workbook":
      session = (payload.tool_input or {}).get("session")
      if not isinstance(session, str) or not session.strip():
        return JSONResponse(
          {"error": "Session token is required", "code": "bad_request"},
          status_code=400,
        )
      try:
        response = await state.relay.switch_active_workbook(session)
      except RuntimeError as exc:
        return _relay_error_response(str(exc))
      return JSONResponse(response)

    try:
      response = await state.relay.execute(payload.tool_name, payload.tool_input or {}, timeout)
    except RuntimeError as exc:
      return _relay_error_response(str(exc))
    except asyncio.TimeoutError:
      return JSONResponse({"error": f"Tool execution timed out after {timeout}s"}, status_code=504)

    return JSONResponse(response)

  @app.get("/api/mcp/events", response_model=None)
  async def mcp_events(
    request: Request,
    secret: str = "",
    session: Optional[str] = None,
    workbook: Optional[str] = None,
    user_id: Optional[str] = None,
  ) -> JSONResponse | StreamingResponse:
    if not _is_valid_mcp_secret(secret, _secret()):
      return JSONResponse({"error": "Unauthorized MCP secret"}, status_code=401)

    try:
      client = await state.relay.register_client(session=session, workbook_name=workbook, user_id=user_id)
    except ValueError as exc:
      return JSONResponse({"error": str(exc)}, status_code=400)

    registry_channel_id = f"excel:{client.gateway_session_id}:{client.client_id}"
    try:
      await state.channel_registry.register(
        Channel(
          channel_id=registry_channel_id,
          channel_type=ChannelType.EXCEL,
          tool_names=state.tool_names,
          metadata={
            "client_id": client.client_id,
            "session": client.gateway_session_id,
            "workbook": client.workbook_name,
          },
        )
      )
    except Exception:
      await state.relay.unregister_client(client.gateway_session_id, client.client_id, immediate=True)
      raise

    async def event_generator() -> AsyncIterator[bytes]:
      last_heartbeat = time.time()
      try:
        while True:
          if await request.is_disconnected():
            break

          now = time.time()
          if now - last_heartbeat >= 15:
            heartbeat = {"type": "heartbeat", "timestamp": int(now)}
            yield f"data: {json_dumps(heartbeat)}\\n\\n".encode("utf-8")
            last_heartbeat = now

          try:
            event = await asyncio.wait_for(client.queue.get(), timeout=1)
          except asyncio.TimeoutError:
            continue

          yield f"data: {json_dumps(event)}\\n\\n".encode("utf-8")

          if event.get("type") == "replaced":
            break
      finally:
        async def _finalize() -> None:
          # Two independent try blocks - see api/main.py version for rationale.
          try:
            await state.relay.unregister_client(client.gateway_session_id, client.client_id)
          except Exception:
            pass
          try:
            await state.channel_registry.unregister(registry_channel_id)
          except Exception:
            pass

        cleanup_task = asyncio.create_task(_finalize())
        state.cleanup_tasks.add(cleanup_task)
        cleanup_task.add_done_callback(state.cleanup_tasks.discard)
        try:
          await asyncio.shield(cleanup_task)
        except asyncio.CancelledError:
          raise

    headers = {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-cache, no-store, must-revalidate",
      "X-Accel-Buffering": "no",
      "Connection": "keep-alive",
    }
    return StreamingResponse(event_generator(), headers=headers)

  @app.post("/api/mcp/tool-result")
  async def mcp_tool_result(payload: McpToolResultRequest) -> JSONResponse:
    if payload.ack and (payload.result is not None or payload.error is not None):
      return JSONResponse({"error": "Ack payload cannot include result or error"}, status_code=400)

    if payload.ack:
      status = await state.relay.ack(payload.request_id, payload.nonce, payload.delivery_id)
      if status == "ok" or status == "completed":
        return JSONResponse({"status": "ok"})
      if status == "not_found":
        return JSONResponse({"error": "Unknown request_id"}, status_code=404)
      if status == "expired":
        return JSONResponse({"error": "Request expired"}, status_code=410)
      if status == "nonce_mismatch":
        return JSONResponse({"error": "Nonce mismatch"}, status_code=409)
      if status == "delivery_id_required":
        return JSONResponse({"error": "delivery_id is required"}, status_code=400)
      if status == "stale_delivery":
        return JSONResponse({"error": "Stale delivery_id"}, status_code=409)
      return JSONResponse({"error": "Invalid request state"}, status_code=409)

    status = await state.relay.complete(
      request_id=payload.request_id,
      nonce=payload.nonce,
      delivery_id=payload.delivery_id,
      result=payload.result,
      error=payload.error,
    )
    if status == "ok" or status == "completed":
      return JSONResponse({"status": "ok"})
    if status == "not_found":
      return JSONResponse({"error": "Unknown request_id"}, status_code=404)
    if status == "expired":
      return JSONResponse({"error": "Request expired"}, status_code=410)
    if status == "nonce_mismatch":
      return JSONResponse({"error": "Nonce mismatch"}, status_code=409)
    if status == "delivery_id_required":
      return JSONResponse({"error": "delivery_id is required"}, status_code=400)
    if status == "stale_delivery":
      return JSONResponse({"error": "Stale delivery_id"}, status_code=409)
    return JSONResponse({"error": "Invalid request state"}, status_code=409)

  return app


app = create_relay_app()


__all__ = [
  "app",
  "create_relay_app",
  "json_dumps",
  "McpRelay",
  "McpExecuteRequest",
  "McpToolResultRequest",
  "ClientState",
  "Channel",
  "ChannelType",
  "ChannelRegistry",
]
