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
  ack: bool = False
  result: Optional[Dict[str, Any]] = None
  error: Optional[Dict[str, Any]] = None


class McpRelay:
  ACK_REPLAY_GRACE_SECONDS = 5
  EXPIRED_TTL_SECONDS = 300

  def __init__(self) -> None:
    self._lock = asyncio.Lock()
    self.inflight: Dict[str, Dict[str, Any]] = {}
    self.active_client_id: Optional[str] = None
    self.active_client_queue: Optional["asyncio.Queue[Dict[str, Any]]"] = None
    self._replay_tasks: Set["asyncio.Task[Any]"] = set()

  async def execute(self, tool_name: str, tool_input: Dict[str, Any], timeout: int) -> Dict[str, Any]:
    request_id = uuid.uuid4().hex
    nonce = os.urandom(12).hex()
    result_queue: "asyncio.Queue[Dict[str, Any]]" = asyncio.Queue(maxsize=1)

    entry = {
      "request_id": request_id,
      "nonce": nonce,
      "tool_name": tool_name,
      "tool_input": tool_input,
      "state": "queued",
      "created_at": time.time(),
      "delivered_at": None,
      "acked_at": None,
      "timeout": timeout,
      "result_queue": result_queue,
    }

    async with self._lock:
      self._prune_expired_locked()
      if self.active_client_queue is None:
        raise RuntimeError("no_frontend")
      self.inflight[request_id] = entry

    delivered = await self._deliver(request_id, replay=False)
    if not delivered:
      async with self._lock:
        self.inflight.pop(request_id, None)
      raise RuntimeError("no_frontend")

    try:
      result_payload = await asyncio.wait_for(result_queue.get(), timeout=timeout)
    except asyncio.TimeoutError:
      async with self._lock:
        existing = self.inflight.get(request_id)
        if existing:
          existing["state"] = "expired"
          existing["expired_at"] = time.time()
      raise

    async with self._lock:
      self.inflight.pop(request_id, None)

    return {
      "request_id": request_id,
      "result": result_payload.get("result"),
      "error": result_payload.get("error"),
    }

  async def register_client(self) -> Tuple[str, "asyncio.Queue[Dict[str, Any]]"]:
    queue: "asyncio.Queue[Dict[str, Any]]" = asyncio.Queue()
    client_id = uuid.uuid4().hex
    replaced_queue: Optional["asyncio.Queue[Dict[str, Any]]"] = None
    replay_ids: List[str] = []
    delayed_replays: List[Tuple[str, float]] = []
    now = time.time()

    async with self._lock:
      self._prune_expired_locked()
      replaced_queue = self.active_client_queue
      self.active_client_id = client_id
      self.active_client_queue = queue

      for request_id, item in self.inflight.items():
        if item.get("state") != "delivered":
          continue
        delivered_at = item.get("delivered_at") or 0
        age = now - float(delivered_at)
        if age >= self.ACK_REPLAY_GRACE_SECONDS:
          replay_ids.append(request_id)
        else:
          delayed_replays.append((request_id, self.ACK_REPLAY_GRACE_SECONDS - age))

    if replaced_queue is not None:
      await replaced_queue.put(
        {
          "type": "replaced",
          "reason": "A newer Excel taskpane connection replaced this stream.",
        }
      )

    for request_id in replay_ids:
      await self._deliver(request_id, replay=True)

    for request_id, delay in delayed_replays:
      self._schedule_delayed_replay(request_id, delay)

    return client_id, queue

  async def unregister_client(self, client_id: str) -> None:
    async with self._lock:
      if self.active_client_id == client_id:
        self.active_client_id = None
        self.active_client_queue = None

  async def ack(self, request_id: str, nonce: str) -> str:
    async with self._lock:
      self._prune_expired_locked()
      entry = self.inflight.get(request_id)
      if not entry:
        return "not_found"
      if entry.get("nonce") != nonce:
        return "nonce_mismatch"
      if entry.get("state") == "expired":
        self.inflight.pop(request_id, None)
        return "expired"
      if entry.get("state") == "completed":
        return "completed"
      entry["state"] = "acked"
      entry["acked_at"] = time.time()
      return "ok"

  async def complete(
    self,
    request_id: str,
    nonce: str,
    result: Optional[Dict[str, Any]],
    error: Optional[Dict[str, Any]],
  ) -> str:
    result_queue: Optional["asyncio.Queue[Dict[str, Any]]"] = None

    async with self._lock:
      self._prune_expired_locked()
      entry = self.inflight.get(request_id)
      if not entry:
        return "not_found"
      if entry.get("nonce") != nonce:
        return "nonce_mismatch"
      if entry.get("state") == "expired":
        self.inflight.pop(request_id, None)
        return "expired"
      if entry.get("state") == "completed":
        return "completed"
      entry["state"] = "completed"
      entry["completed_at"] = time.time()
      result_queue = entry.get("result_queue")

    if result_queue is not None:
      await result_queue.put({"result": result, "error": error})

    return "ok"

  async def _deliver(self, request_id: str, replay: bool) -> bool:
    async with self._lock:
      entry = self.inflight.get(request_id)
      queue = self.active_client_queue
      if not entry or queue is None:
        return False
      entry["state"] = "delivered"
      entry["delivered_at"] = time.time()
      payload = {
        "type": "mcp_tool_request",
        "request_id": request_id,
        "nonce": entry.get("nonce"),
        "tool_name": entry.get("tool_name"),
        "tool_input": entry.get("tool_input"),
        "replay": replay,
      }

    await queue.put(payload)
    return True

  def _schedule_delayed_replay(self, request_id: str, delay: float) -> None:
    task: Optional["asyncio.Task[Any]"] = None

    async def _runner() -> None:
      try:
        await asyncio.sleep(max(delay, 0))
        async with self._lock:
          entry = self.inflight.get(request_id)
          if not entry or entry.get("state") != "delivered":
            return
        await self._deliver(request_id, replay=True)
      finally:
        if task is not None:
          self._replay_tasks.discard(task)

    task = asyncio.create_task(_runner())
    self._replay_tasks.add(task)

  def _prune_expired_locked(self) -> None:
    now = time.time()
    to_delete: List[str] = []
    for request_id, item in self.inflight.items():
      if item.get("state") != "expired":
        continue
      expired_at = float(item.get("expired_at") or now)
      if now - expired_at >= self.EXPIRED_TTL_SECONDS:
        to_delete.append(request_id)
    for request_id in to_delete:
      self.inflight.pop(request_id, None)


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
    try:
      response = await state.relay.execute(payload.tool_name, payload.tool_input or {}, timeout)
    except RuntimeError as exc:
      if str(exc) == "no_frontend":
        return JSONResponse({"error": "No active Excel taskpane connection"}, status_code=503)
      return JSONResponse({"error": f"MCP relay failed: {exc}"}, status_code=500)
    except asyncio.TimeoutError:
      return JSONResponse({"error": f"Tool execution timed out after {timeout}s"}, status_code=504)

    return JSONResponse(response)

  @app.get("/api/mcp/events", response_model=None)
  async def mcp_events(request: Request, secret: str = "") -> JSONResponse | StreamingResponse:
    if not _is_valid_mcp_secret(secret, _secret()):
      return JSONResponse({"error": "Unauthorized MCP secret"}, status_code=401)

    client_id, queue = await state.relay.register_client()
    registry_channel_id = f"excel:{client_id}"
    try:
      await state.channel_registry.register(
        Channel(
          channel_id=registry_channel_id,
          channel_type=ChannelType.EXCEL,
          tool_names=state.tool_names,
          metadata={"client_id": client_id},
        )
      )
    except Exception:
      await state.relay.unregister_client(client_id)
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
            event = await asyncio.wait_for(queue.get(), timeout=1)
          except asyncio.TimeoutError:
            continue

          yield f"data: {json_dumps(event)}\\n\\n".encode("utf-8")

          if event.get("type") == "replaced":
            break
      finally:
        await state.relay.unregister_client(client_id)
        await state.channel_registry.unregister(registry_channel_id)

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
      status = await state.relay.ack(payload.request_id, payload.nonce)
      if status == "ok" or status == "completed":
        return JSONResponse({"status": "ok"})
      if status == "not_found":
        return JSONResponse({"error": "Unknown request_id"}, status_code=404)
      if status == "expired":
        return JSONResponse({"error": "Request expired"}, status_code=410)
      if status == "nonce_mismatch":
        return JSONResponse({"error": "Nonce mismatch"}, status_code=409)
      return JSONResponse({"error": "Invalid request state"}, status_code=409)

    status = await state.relay.complete(
      request_id=payload.request_id,
      nonce=payload.nonce,
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
  "Channel",
  "ChannelType",
  "ChannelRegistry",
]
