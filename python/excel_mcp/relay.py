from __future__ import annotations

import asyncio
import math
import os
import time
import uuid
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, AsyncIterator, Awaitable, Callable, Dict, List, Literal, Optional, Set, Tuple

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pydantic import BaseModel, Field

from .tool_registry import get_tool_specs


_CHAT_RELAY_REQUEST_ID_TOOL_INPUT_KEY = "_relay_request_id"


class McpExecuteRequest(BaseModel):
  kind: Literal["tool", "chat"] = "tool"
  tool_name: str
  tool_input: Dict[str, Any] = Field(default_factory=dict)
  request_id: Optional[str] = None

  def model_post_init(self, __context: Any) -> None:
    if self.kind != "chat" or self.request_id is None:
      return
    tool_input = dict(self.tool_input or {})
    tool_input[_CHAT_RELAY_REQUEST_ID_TOOL_INPUT_KEY] = self.request_id
    self.tool_input = tool_input


class McpToolResultRequest(BaseModel):
  request_id: str
  nonce: str
  delivery_id: Optional[str] = None
  ack: bool = False
  result: Optional[Dict[str, Any]] = None
  error: Optional[Dict[str, Any]] = None


@dataclass(frozen=True)
class RelayAuthContext:
  session_id: str
  user_id: str
  channel: Optional[str] = None


class RelayAuthError(RuntimeError):
  def __init__(self, message: str, *, status_code: int = 401) -> None:
    self.status_code = status_code
    super().__init__(message)


RelayAuthenticator = Callable[[Request], Awaitable[RelayAuthContext]]


_TOOL_TIMEOUT = int(os.getenv("EXCEL_MCP_TOOL_TIMEOUT", "180"))
RELAY_RECONNECT_GRACE_S = int(os.getenv("RELAY_RECONNECT_GRACE_S", str(_TOOL_TIMEOUT + 30)))
CHAT_RELAY_DEV_ENV = "EXCEL_CHAT_RELAY_DEV"


def chat_relay_dev_enabled() -> bool:
  return os.getenv(CHAT_RELAY_DEV_ENV, "").strip() == "1"


def chat_relay_timeout_seconds(tool_input: Dict[str, Any], default: int = 600) -> int:
  raw_timeout = tool_input.get("timeout_s", default)
  try:
    timeout = float(raw_timeout)
  except (TypeError, ValueError):
    return default
  if not math.isfinite(timeout) or timeout <= 0:
    return default
  return max(1, int(math.ceil(timeout)))


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
    self._user_active_workbook: Dict[str, str] = {}
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
    kind: str = "tool",
    request_id: Optional[str] = None,
  ) -> Dict[str, Any]:
    request_id, result_future = await self._enqueue(
      tool_name,
      tool_input,
      timeout,
      target_session=target_session,
      gateway_session_id=gateway_session_id,
      user_id=user_id,
      kind=kind,
      request_id=request_id if kind == "chat" else None,
    )

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

  async def submit(
    self,
    tool_name: str,
    tool_input: Dict[str, Any],
    timeout: int,
    target_session: Optional[str] = None,
    *,
    gateway_session_id: Optional[str] = None,
    user_id: Optional[str] = None,
    kind: str = "tool",
    request_id: Optional[str] = None,
  ) -> Dict[str, Any]:
    request_id, _ = await self._enqueue(
      tool_name,
      tool_input,
      timeout,
      target_session=target_session,
      gateway_session_id=gateway_session_id,
      user_id=user_id,
      kind=kind,
      request_id=request_id if kind == "chat" else None,
    )
    return {"request_id": request_id}

  async def result(
    self,
    request_id: str,
    *,
    gateway_session_id: Optional[str] = None,
    user_id: Optional[str] = None,
  ) -> Tuple[str, Dict[str, Any]]:
    async with self._lock:
      self._prune_expired_locked()
      entry = self._inflight.get(request_id)
      if not entry:
        return "not_found", {
          "state": "timeout",
          "error": {"message": "Request not found or no longer retained"},
        }
      ownership_status = self._validate_result_owner_locked(entry, gateway_session_id, user_id)
      if ownership_status != "ok":
        return ownership_status, {}

      now = time.time()
      entry_state = entry.get("state")
      if entry_state in {"queued", "delivered", "acked"}:
        timeout = float(entry.get("timeout") or 0)
        created_at = float(entry.get("created_at") or now)
        if timeout > 0 and now - created_at >= timeout:
          entry["state"] = "expired"
          entry["expired_at"] = now
          entry_state = "expired"

      if entry_state == "completed":
        error = entry.get("error")
        if error is not None:
          return "ok", {"state": "failed", "error": error}
        return "ok", {"state": "done", "result": entry.get("result")}

      if entry_state == "expired":
        return "ok", {
          "state": "timeout",
          "error": {"message": "Request timed out"},
        }

      if entry_state == "orphaned":
        return "ok", {
          "state": "failed",
          "error": {"message": "Target workbook session disconnected"},
        }

      return "ok", {"state": "pending"}

  async def _enqueue(
    self,
    tool_name: str,
    tool_input: Dict[str, Any],
    timeout: int,
    target_session: Optional[str] = None,
    *,
    gateway_session_id: Optional[str] = None,
    user_id: Optional[str] = None,
    kind: str = "tool",
    request_id: Optional[str] = None,
  ) -> Tuple[str, "asyncio.Future[Dict[str, Any]]"]:
    if kind not in {"tool", "chat"}:
      raise RuntimeError("invalid_kind")

    nonce = os.urandom(12).hex()
    result_future: "asyncio.Future[Dict[str, Any]]" = asyncio.get_running_loop().create_future()
    if kind == "chat":
      request_id_from_payload = tool_input.get(_CHAT_RELAY_REQUEST_ID_TOOL_INPUT_KEY)
      if request_id is None:
        request_id = request_id_from_payload
    if request_id is not None:
      if not isinstance(request_id, str) or not request_id.strip():
        raise RuntimeError("invalid_request_id")
      request_id = request_id.strip()
    else:
      request_id = uuid.uuid4().hex
    tool_payload = dict(tool_input)
    if kind == "chat":
      tool_payload.pop(_CHAT_RELAY_REQUEST_ID_TOOL_INPUT_KEY, None)
    requested_session = target_session or tool_payload.pop("_workbook", None)

    async with self._lock:
      self._prune_expired_locked()
      if requested_session is not None and not isinstance(requested_session, str):
        raise RuntimeError("unknown_session")
      client = self._resolve_client_locked(gateway_session_id, requested_session, user_id)
      if client is None:
        if requested_session is not None:
          raise RuntimeError("unknown_session")
        raise RuntimeError("no_active_session")
      if user_id is not None and client.user_id != user_id:
        raise RuntimeError("session_disconnected")
      if self._client_detached_past_grace_locked(client):
        raise RuntimeError("session_disconnected")
      if kind == "chat" and request_id in self._inflight:
        raise RuntimeError("request_id_inflight")

      chat_payload = tool_payload if kind == "chat" else {}
      entry = {
        "request_id": request_id,
        "nonce": nonce,
        "kind": kind,
        "tool_name": tool_name,
        "tool_input": tool_payload,
        "state": "queued",
        "created_at": time.time(),
        "delivered_at": None,
        "acked_at": None,
        "timeout": timeout,
        "session_id": client.gateway_session_id,
        "user_id": client.user_id,
        "owner_session_id": gateway_session_id,
        "owner_user_id": user_id,
        "workbook_session": client.workbook_session,
        "target_session": client.gateway_session_id,
        "delivery_client_id": None,
        "future": result_future,
        "result": None,
        "error": None,
      }
      if kind == "chat":
        entry["delegation_id"] = (
          chat_payload.get("delegation_id") if isinstance(chat_payload, dict) else None
        )
      self._inflight[request_id] = entry

    delivered = await self._deliver(request_id, replay=False)
    if not delivered:
      async with self._lock:
        self._inflight.pop(request_id, None)
      raise RuntimeError("session_disconnected")

    return request_id, result_future

  async def list_workbooks(
    self,
    gateway_session_id: Optional[str] = None,
    user_id: Optional[str] = None,
  ) -> List[Dict[str, Any]]:
    async with self._lock:
      self._prune_expired_locked()
      clients = self._clients.values()
      if user_id is not None:
        clients = [client for client in clients if client.user_id == user_id]
      workbooks = [
        {
          "name": client.workbook_name,
          "session": client.workbook_session,
          "gateway_session_id": client.gateway_session_id,
          "active": self._client_is_active_locked(client, user_id),
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
      client = self._resolve_client_locked(gateway_session_id, session, user_id)
      if client is None:
        raise RuntimeError("unknown_session")
      if user_id is not None and client.user_id != user_id:
        raise RuntimeError("unknown_session")
      self._active_workbook[client.gateway_session_id] = client.workbook_session
      if client.user_id is not None:
        self._user_active_workbook[client.user_id] = client.gateway_session_id

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
      if client.user_id is not None:
        existing_active = self._user_active_workbook.get(client.user_id)
        existing_active_client = self._clients.get(existing_active) if existing_active is not None else None
        if existing_active_client is None or existing_active_client.detached:
          self._user_active_workbook[client.user_id] = client.gateway_session_id

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
      entry["result"] = result
      entry["error"] = error
      entry["state"] = "completed"
      entry["completed_at"] = time.time()
      result_future = entry.get("future")

    if result_future is not None and not result_future.done():
      result_future.set_result({"result": result, "error": error})

    return "ok"

  async def status(
    self,
    gateway_session_id: Optional[str],
    user_id: Optional[str] = None,
  ) -> Dict[str, Any]:
    async with self._lock:
      self._prune_expired_locked()
      status_session_id = gateway_session_id
      client = self._clients.get(gateway_session_id) if gateway_session_id is not None else None
      if client is not None and user_id is not None and client.user_id != user_id:
        client = None
      if client is None and user_id is not None:
        user_session = self._user_active_workbook.get(user_id)
        user_client = self._clients.get(user_session) if user_session is not None else None
        if user_client is not None and user_client.user_id == user_id:
          client = user_client
          status_session_id = user_client.gateway_session_id
      inflight_count = sum(
        1
        for item in self._inflight.values()
        if item.get("session_id") == status_session_id
        and (user_id is None or item.get("user_id") == user_id)
        and item.get("state") not in {"completed", "expired"}
      )
      return {
        "session_id": status_session_id,
        "connected": bool(client and not client.detached),
        "detached": bool(client and client.detached),
        "workbook_session": client.workbook_session if client else None,
        "workbook_name": client.workbook_name if client else None,
        "active_workbook": self._active_workbook.get(status_session_id) if status_session_id else None,
        "tools": self._session_tools.get(status_session_id, []) if status_session_id else [],
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
    if entry.get("kind") == "chat":
      tool_input = entry.get("tool_input")
      chat_payload = tool_input if isinstance(tool_input, dict) else {}
      payload = {
        "type": "mcp_chat_request",
        "request_id": request_id,
        "nonce": entry.get("nonce"),
        "delivery_id": delivery_id,
        "text": chat_payload.get("text"),
        "force_compaction": chat_payload.get("force_compaction", False),
        "delegation_id": entry.get("delegation_id"),
        "replay": replay,
      }
      if "seed_history" in chat_payload:
        payload["seed_history"] = chat_payload.get("seed_history")
      if "model" in chat_payload:
        payload["model"] = chat_payload.get("model")
      return payload
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

  def _validate_result_owner_locked(
    self,
    entry: Dict[str, Any],
    gateway_session_id: Optional[str],
    user_id: Optional[str],
  ) -> str:
    if gateway_session_id is not None and entry.get("owner_session_id") != gateway_session_id:
      return "session_mismatch"
    if user_id is not None and entry.get("owner_user_id") != user_id:
      return "user_mismatch"
    return "ok"

  def _resolve_client_locked(
    self,
    gateway_session_id: Optional[str],
    requested_session: Optional[str],
    user_id: Optional[str] = None,
  ) -> Optional[ClientState]:
    if requested_session:
      client = self._clients.get(requested_session)
      if client is None:
        for candidate in self._clients.values():
          if candidate.workbook_session == requested_session:
            client = candidate
            break
      if client is None:
        return None
      if user_id is not None and client.user_id != user_id:
        return None
      return client

    if gateway_session_id is not None:
      client = self._clients.get(gateway_session_id)
      if client is not None:
        if user_id is not None and client.user_id != user_id:
          return None
        return client

    if user_id is not None:
      user_session = self._user_active_workbook.get(user_id)
      if user_session is not None:
        client = self._clients.get(user_session)
        if client is not None and client.user_id == user_id:
          return client

    return None

  def _client_detached_past_grace_locked(self, client: ClientState) -> bool:
    deadline = client.detach_grace_deadline
    return deadline is not None and time.time() >= deadline

  def _client_is_active_locked(self, client: ClientState, user_id: Optional[str]) -> bool:
    if user_id is not None:
      user_session = self._user_active_workbook.get(user_id)
      if user_session is not None:
        return client.gateway_session_id == user_session
    return self._active_workbook.get(client.gateway_session_id) == client.workbook_session

  def _replacement_user_session_locked(self, user_id: str) -> Optional[str]:
    detached_replacement: Optional[str] = None
    for client in self._clients.values():
      if client.user_id != user_id:
        continue
      if not client.detached:
        return client.gateway_session_id
      if detached_replacement is None:
        detached_replacement = client.gateway_session_id
    return detached_replacement

  def _remove_client_locked(self, session: str, *, orphan_inflight: bool = False) -> None:
    removed_client = self._clients.pop(session, None)
    self._active_workbook.pop(session, None)
    self._session_tools.pop(session, None)
    if removed_client is not None and removed_client.user_id is not None:
      if self._user_active_workbook.get(removed_client.user_id) == session:
        replacement = self._replacement_user_session_locked(removed_client.user_id)
        if replacement is not None:
          self._user_active_workbook[removed_client.user_id] = replacement
        else:
          self._user_active_workbook.pop(removed_client.user_id, None)
    if orphan_inflight:
      for item in self._inflight.values():
        if item.get("session_id") == session and item.get("state") in {"queued", "delivered", "acked"}:
          item["state"] = "orphaned"
          item["orphaned_at"] = time.time()

  def _prune_expired_locked(self) -> None:
    now = time.time()
    to_delete: List[str] = []
    for request_id, item in self._inflight.items():
      state = item.get("state")
      if state not in {"completed", "expired"}:
        continue
      reference_at = item.get("completed_at") if state == "completed" else item.get("expired_at")
      reference_at = float(reference_at or now)
      if now - reference_at >= self.EXPIRED_TTL_SECONDS:
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


def create_relay_app(
  authenticate_request: RelayAuthenticator | None = None,
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
    allow_headers=["Authorization", "Content-Type"],
  )

  async def _authenticate(request: Request) -> RelayAuthContext:
    if authenticate_request is None:
      raise RelayAuthError("Relay auth is not configured")
    context = await authenticate_request(request)
    session_id = str(context.session_id or "").strip()
    user_id = str(context.user_id or "").strip()
    if not session_id:
      raise RelayAuthError("Relay auth session_id is required")
    if not user_id:
      raise RelayAuthError("Relay auth user_id is required")
    return RelayAuthContext(session_id=session_id, user_id=user_id, channel=context.channel)

  def _auth_error_response(exc: RelayAuthError) -> JSONResponse:
    return JSONResponse({"error": str(exc)}, status_code=exc.status_code)

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
    try:
      await _authenticate(request)
    except RelayAuthError as exc:
      return _auth_error_response(exc)
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
    try:
      auth_context = await _authenticate(request)
    except RelayAuthError as exc:
      return _auth_error_response(exc)

    if payload.kind == "chat":
      if not chat_relay_dev_enabled():
        return JSONResponse({"error": "Chat relay is disabled"}, status_code=403)
      timeout = chat_relay_timeout_seconds(payload.tool_input or {})
      try:
        response = await state.relay.submit(
          payload.tool_name,
          payload.tool_input or {},
          timeout,
          gateway_session_id=auth_context.session_id,
          user_id=auth_context.user_id,
          kind="chat",
          request_id=payload.request_id,
        )
      except RuntimeError as exc:
        return _relay_error_response(str(exc))
      return JSONResponse({"request_id": response["request_id"]})

    timeout = int(os.getenv("EXCEL_MCP_TOOL_TIMEOUT", "60"))
    if payload.tool_name == "list_workbooks":
      return JSONResponse({
        "workbooks": await state.relay.list_workbooks(
          gateway_session_id=auth_context.session_id,
          user_id=auth_context.user_id,
        )
      })
    if payload.tool_name == "switch_active_workbook":
      session = (payload.tool_input or {}).get("session")
      if not isinstance(session, str) or not session.strip():
        return JSONResponse(
          {"error": "Session token is required", "code": "bad_request"},
          status_code=400,
        )
      try:
        response = await state.relay.switch_active_workbook(
          session,
          gateway_session_id=auth_context.session_id,
          user_id=auth_context.user_id,
        )
      except RuntimeError as exc:
        return _relay_error_response(str(exc))
      return JSONResponse(response)

    try:
      response = await state.relay.execute(
        payload.tool_name,
        payload.tool_input or {},
        timeout,
        gateway_session_id=auth_context.session_id,
        user_id=auth_context.user_id,
        kind="tool",
      )
    except RuntimeError as exc:
      return _relay_error_response(str(exc))
    except asyncio.TimeoutError:
      return JSONResponse({"error": f"Tool execution timed out after {timeout}s"}, status_code=504)

    return JSONResponse(response)

  @app.get("/api/mcp/result/{request_id}")
  async def mcp_result(request: Request, request_id: str) -> JSONResponse:
    try:
      auth_context = await _authenticate(request)
    except RelayAuthError as exc:
      return _auth_error_response(exc)

    status, response = await state.relay.result(
      request_id,
      gateway_session_id=auth_context.session_id,
      user_id=auth_context.user_id,
    )
    if status == "ok":
      return JSONResponse(response)
    if status == "not_found":
      return JSONResponse(response, status_code=404)
    if status in {"session_mismatch", "user_mismatch"}:
      return JSONResponse({"error": "Request does not belong to this session"}, status_code=403)
    return JSONResponse({"error": "Invalid request state"}, status_code=409)

  @app.get("/api/mcp/events", response_model=None)
  async def mcp_events(
    request: Request,
    session: Optional[str] = None,
    workbook: Optional[str] = None,
  ) -> JSONResponse | StreamingResponse:
    try:
      auth_context = await _authenticate(request)
    except RelayAuthError as exc:
      return _auth_error_response(exc)
    workbook_session = session.strip() if isinstance(session, str) and session.strip() else auth_context.session_id

    try:
      client = await state.relay.register_client(
        session=workbook_session,
        workbook_name=workbook,
        gateway_session_id=auth_context.session_id,
        user_id=auth_context.user_id,
        workbook_session=workbook_session,
      )
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
            "workbook_session": client.workbook_session,
            "workbook": client.workbook_name,
            "user_id": client.user_id,
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
  async def mcp_tool_result(request: Request, payload: McpToolResultRequest) -> JSONResponse:
    try:
      auth_context = await _authenticate(request)
    except RelayAuthError as exc:
      return _auth_error_response(exc)

    if payload.ack and (payload.result is not None or payload.error is not None):
      return JSONResponse({"error": "Ack payload cannot include result or error"}, status_code=400)

    if payload.ack:
      status = await state.relay.ack(
        payload.request_id,
        payload.nonce,
        payload.delivery_id,
        gateway_session_id=auth_context.session_id,
        user_id=auth_context.user_id,
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
      if status in {"session_mismatch", "user_mismatch"}:
        return JSONResponse({"error": "Request does not belong to this session"}, status_code=403)
      return JSONResponse({"error": "Invalid request state"}, status_code=409)

    status = await state.relay.complete(
      request_id=payload.request_id,
      nonce=payload.nonce,
      delivery_id=payload.delivery_id,
      result=payload.result,
      error=payload.error,
      gateway_session_id=auth_context.session_id,
      user_id=auth_context.user_id,
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
    if status in {"session_mismatch", "user_mismatch"}:
      return JSONResponse({"error": "Request does not belong to this session"}, status_code=403)
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
  "RelayAuthContext",
  "RelayAuthError",
  "ClientState",
  "Channel",
  "ChannelType",
  "ChannelRegistry",
  "CHAT_RELAY_DEV_ENV",
  "chat_relay_dev_enabled",
  "chat_relay_timeout_seconds",
]
