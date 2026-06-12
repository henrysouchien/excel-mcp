#!/usr/bin/env python3
"""Excel MCP server that proxies tool calls to the relay backend."""

# CRITICAL: Redirect stdout to stderr BEFORE any imports.
# MCP uses stdout for JSON-RPC traffic, so logs/prints must go to stderr.
import sys
_real_stdout = sys.stdout
sys.stdout = sys.stderr

import asyncio
import json
import os
import signal
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict
from urllib.parse import urlparse, urlunparse

import httpx

try:
  from fastmcp import FastMCP
  from fastmcp.tools.tool import Tool, ToolResult
  from mcp.types import TextContent
  MCP_RUNTIME_IMPORT_ERROR: Exception | None = None
except Exception as exc:
  MCP_RUNTIME_IMPORT_ERROR = exc

  @dataclass
  class TextContent:
    type: str
    text: str

  class ToolResult:
    def __init__(
      self,
      *,
      structured_content: Any | None = None,
      content: Any | None = None,
    ) -> None:
      self.structured_content = structured_content
      self.content = content

  class Tool:
    def __init__(self, **kwargs: Any) -> None:
      self.parameters = kwargs.get("parameters", {})

  class FastMCP:
    def __init__(self, name: str, instructions: str = "") -> None:
      self.name = name
      self.instructions = instructions

    def tool(self):  # type: ignore[no-untyped-def]
      def _decorator(func):
        return func

      return _decorator

    def add_tool(self, tool: Any) -> None:
      _ = tool

    def run(self) -> None:
      raise RuntimeError(f"MCP runtime unavailable: {MCP_RUNTIME_IMPORT_ERROR}")

# Restore stdout for MCP communication.
sys.stdout = _real_stdout

from .tool_registry import get_tool_specs

MCP_NAME = "excel-addin"
BACKEND_URL = os.getenv("EXCEL_MCP_BACKEND_URL", "https://localhost:8000/api/mcp/execute")
BACKEND_BASE_URL = os.getenv("EXCEL_MCP_BACKEND_BASE_URL", "").strip().rstrip("/")
USER_API_KEY = os.getenv("EXCEL_MCP_API_KEY", "").strip()
_TLS_VERIFY_RAW = os.getenv("EXCEL_MCP_TLS_VERIFY", "").strip().lower()
GATEWAY_BASE_URL = os.getenv("EXCEL_MCP_GATEWAY_BASE_URL", "").strip().rstrip("/")
DEFAULT_TIMEOUT_SECONDS = int(os.getenv("EXCEL_MCP_TOOL_TIMEOUT", "60"))
_TRUTHY_ENV_VALUES = {"1", "true", "yes", "on"}
CHAT_RELAY_TOOL_NAME = "send_chat_message"
CHAT_RELAY_DEV_ENV = "EXCEL_CHAT_RELAY_DEV"
CHAT_RELAY_DEFAULT_TIMEOUT_SECONDS = 600
CHAT_RELAY_HTTP_TIMEOUT_SECONDS = 10
CHAT_RELAY_POLL_INTERVAL_SECONDS = float(os.getenv("EXCEL_CHAT_RELAY_POLL_INTERVAL", "1"))

mcp = FastMCP(
  MCP_NAME,
  instructions="Excel workbook read/write tools via the local Excel add-in taskpane.",
)


def _derive_gateway_base_url(backend_url: str) -> str:
  parsed = urlparse(backend_url)
  if not (parsed.scheme and parsed.netloc):
    return backend_url.strip().rstrip("/")
  segments = [segment for segment in parsed.path.split("/") if segment]
  if len(segments) >= 3 and segments[-3:-1] == ["api", "mcp"]:
    segments = segments[:-3]
  new_path = "/" + "/".join(segments) if segments else ""
  return urlunparse((parsed.scheme, parsed.netloc, new_path, "", "", "")).rstrip("/")


_derive_backend_base_url = _derive_gateway_base_url


if not BACKEND_BASE_URL:
  BACKEND_BASE_URL = _derive_gateway_base_url(BACKEND_URL)

if not GATEWAY_BASE_URL:
  GATEWAY_BASE_URL = _derive_gateway_base_url(BACKEND_URL)


def _chat_relay_dev_enabled() -> bool:
  return os.getenv(CHAT_RELAY_DEV_ENV, "").strip() == "1"


def _pid_file_path() -> Path:
  override = os.getenv("EXCEL_MCP_PID_FILE", "").strip()
  if override:
    return Path(override)
  return Path.cwd() / ".excel_mcp_server.pid"


def _singleton_pid_kill_enabled() -> bool:
  return os.getenv("EXCEL_MCP_SINGLETON", "").strip().lower() in _TRUTHY_ENV_VALUES


def _write_current_pid() -> None:
  _pid_file_path().write_text(str(os.getpid()))


def _kill_previous_instance() -> None:
  """Kill any previous MCP server instance using a PID file."""
  pid_file = _pid_file_path()
  if pid_file.exists():
    try:
      old_pid = int(pid_file.read_text().strip())
      if old_pid != os.getpid():
        os.kill(old_pid, signal.SIGTERM)
        print(f"Killed previous MCP server (PID {old_pid})", file=sys.stderr)
    except (ValueError, ProcessLookupError, PermissionError):
      pass
  _write_current_pid()


def _prepare_stdio_instance() -> None:
  """Prepare process state for a stdio MCP server.

  Multiple Claude sessions may legitimately spawn independent stdio proxy
  processes. Keep singleton termination opt-in so a new session does not
  disconnect an existing session.
  """
  if _singleton_pid_kill_enabled():
    _kill_previous_instance()
    return
  _write_current_pid()


def _call_backend(
  tool_name: str,
  tool_input: Dict[str, Any],
  timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS,
) -> Any:
  from ._gateway_session import default_tls_verify, get_session_token, invalidate

  tls_verify = (
    _TLS_VERIFY_RAW in _TRUTHY_ENV_VALUES
    if _TLS_VERIFY_RAW
    else default_tls_verify(GATEWAY_BASE_URL)
  )

  payload = {
    "tool_name": tool_name,
    "tool_input": tool_input,
  }
  headers = {
    "Content-Type": "application/json",
  }
  if USER_API_KEY:
    token = get_session_token(USER_API_KEY, gateway_base_url=GATEWAY_BASE_URL, tls_verify=tls_verify)
    headers["Authorization"] = f"Bearer {token}"
  else:
    raise RuntimeError(
      "EXCEL_MCP_API_KEY is required; set it to a channel='mcp' GATEWAY_USER_KEYS key."
    )

  try:
    with httpx.Client(verify=tls_verify, timeout=timeout_seconds) as client:
      response = client.post(BACKEND_URL, json=payload, headers=headers)
  except httpx.HTTPError as exc:
    raise RuntimeError(f"Backend request failed: {exc}") from exc

  if response.status_code == 401 and USER_API_KEY:
    invalidate(USER_API_KEY, GATEWAY_BASE_URL)
    token = get_session_token(USER_API_KEY, gateway_base_url=GATEWAY_BASE_URL, tls_verify=tls_verify)
    headers["Authorization"] = f"Bearer {token}"
    try:
      with httpx.Client(verify=tls_verify, timeout=timeout_seconds) as client:
        response = client.post(BACKEND_URL, json=payload, headers=headers)
    except httpx.HTTPError as exc:
      raise RuntimeError(f"Backend request failed after re-auth: {exc}") from exc

  if response.status_code >= 400:
    try:
      body = response.json()
    except Exception:
      body = {"error": response.text or f"HTTP {response.status_code}"}
    error_text = body.get("error") if isinstance(body, dict) else None
    if isinstance(error_text, str):
      lowered = error_text.lower()
      if "no active excel taskpane connection" in lowered or "no_frontend" in lowered:
        raise RuntimeError(
          f"Excel taskpane is not connected. Tool '{tool_name}' requires an active Excel session. "
          "Open Excel and load the add-in, then try again."
        )
    raise RuntimeError(error_text or f"Backend request failed ({response.status_code})")

  try:
    data = response.json()
  except Exception as exc:
    raise RuntimeError("Backend returned invalid JSON") from exc

  if not isinstance(data, dict):
    raise RuntimeError("Backend response must be an object")

  if data.get("error"):
    err = data["error"]
    if isinstance(err, dict):
      message = err.get("message") or err.get("error") or str(err)
    else:
      message = str(err)
    lowered = message.lower()
    if "no active excel taskpane connection" in lowered or "no_frontend" in lowered:
      raise RuntimeError(
        f"Excel taskpane is not connected. Tool '{tool_name}' requires an active Excel session. "
        "Open Excel and load the add-in, then try again."
      )
    raise RuntimeError(message)

  if "result" in data:
    return data.get("result")
  return data


def _authorized_backend_json(
  method: str,
  url: str,
  *,
  payload: Dict[str, Any] | None = None,
  timeout_seconds: float = CHAT_RELAY_HTTP_TIMEOUT_SECONDS,
) -> Dict[str, Any]:
  from ._gateway_session import default_tls_verify, get_session_token, invalidate

  tls_verify = (
    _TLS_VERIFY_RAW in _TRUTHY_ENV_VALUES
    if _TLS_VERIFY_RAW
    else default_tls_verify(GATEWAY_BASE_URL)
  )
  if not USER_API_KEY:
    raise RuntimeError(
      "EXCEL_MCP_API_KEY is required; set it to a channel='mcp' GATEWAY_USER_KEYS key."
    )

  headers = {"Content-Type": "application/json"}

  def _send(token: str) -> httpx.Response:
    headers["Authorization"] = f"Bearer {token}"
    with httpx.Client(verify=tls_verify, timeout=timeout_seconds) as client:
      if method.upper() == "GET":
        return client.get(url, headers=headers)
      if method.upper() == "POST":
        return client.post(url, json=payload or {}, headers=headers)
    raise RuntimeError(f"Unsupported backend method: {method}")

  try:
    token = get_session_token(USER_API_KEY, gateway_base_url=GATEWAY_BASE_URL, tls_verify=tls_verify)
    response = _send(token)
  except httpx.HTTPError as exc:
    raise RuntimeError(f"Backend request failed: {exc}") from exc

  if response.status_code == 401:
    invalidate(USER_API_KEY, GATEWAY_BASE_URL)
    try:
      token = get_session_token(USER_API_KEY, gateway_base_url=GATEWAY_BASE_URL, tls_verify=tls_verify)
      response = _send(token)
    except httpx.HTTPError as exc:
      raise RuntimeError(f"Backend request failed after re-auth: {exc}") from exc

  if response.status_code >= 400:
    try:
      body = response.json()
    except Exception:
      body = {"error": response.text or f"HTTP {response.status_code}"}
    error_text = body.get("error") if isinstance(body, dict) else None
    if isinstance(error_text, dict):
      error_text = error_text.get("message") or json.dumps(error_text, sort_keys=True)
    raise RuntimeError(str(error_text or f"Backend request failed ({response.status_code})"))

  try:
    data = response.json()
  except Exception as exc:
    raise RuntimeError("Backend returned invalid JSON") from exc

  if not isinstance(data, dict):
    raise RuntimeError("Backend response must be an object")
  return data


def _chat_result_url(request_id: str) -> str:
  return f"{BACKEND_BASE_URL}/api/mcp/result/{request_id}"


def _coerce_chat_timeout(value: Any) -> float:
  if value is None:
    return float(CHAT_RELAY_DEFAULT_TIMEOUT_SECONDS)
  try:
    timeout_s = float(value)
  except (TypeError, ValueError):
    return float(CHAT_RELAY_DEFAULT_TIMEOUT_SECONDS)
  if timeout_s <= 0:
    return float(CHAT_RELAY_DEFAULT_TIMEOUT_SECONDS)
  return timeout_s


def _coerce_seed_history(value: Any) -> list[Dict[str, str]]:
  if value is None:
    return []
  if not isinstance(value, list):
    raise RuntimeError("seed_history must be an array")

  messages: list[Dict[str, str]] = []
  for index, item in enumerate(value):
    if not isinstance(item, dict):
      raise RuntimeError(f"seed_history[{index}] must be an object")
    role = item.get("role")
    content = item.get("content")
    if role not in {"user", "assistant"}:
      raise RuntimeError(f"seed_history[{index}].role must be 'user' or 'assistant'")
    if not isinstance(content, str):
      raise RuntimeError(f"seed_history[{index}].content must be a string")
    messages.append({"role": role, "content": content})
  return messages


def _submit_chat_message(
  text: str,
  force_compaction: bool,
  timeout_s: float,
  seed_history: list[Dict[str, str]] | None = None,
  model: str | None = None,
  approve_tool_classes: list[str] | None = None,
  approval_window_seconds: Any = None,
  workbook: str | None = None,
) -> Dict[str, str]:
  tool_input = {
    "text": text,
    "force_compaction": force_compaction,
    "timeout_s": timeout_s,
  }
  if seed_history:
    tool_input["seed_history"] = seed_history
  if model:
    tool_input["model"] = model
  if approve_tool_classes:
    tool_input["approve_tool_classes"] = approve_tool_classes
  if approval_window_seconds is not None:
    tool_input["approval_window_seconds"] = approval_window_seconds
  if workbook:
    tool_input["workbook"] = workbook
  payload = {
    "kind": "chat",
    "tool_name": CHAT_RELAY_TOOL_NAME,
    "tool_input": tool_input,
  }
  response = _authorized_backend_json(
    "POST",
    BACKEND_URL,
    payload=payload,
    timeout_seconds=CHAT_RELAY_HTTP_TIMEOUT_SECONDS,
  )
  request_id = response.get("request_id")
  if not isinstance(request_id, str) or not request_id:
    raise RuntimeError("Chat relay submit response did not include request_id")
  submitted = {"request_id": request_id}
  delegation_id = response.get("delegation_id")
  if isinstance(delegation_id, str) and delegation_id:
    submitted["delegation_id"] = delegation_id
  return submitted


def _poll_chat_message(request_id: str, timeout_seconds: float = CHAT_RELAY_HTTP_TIMEOUT_SECONDS) -> Dict[str, Any]:
  response = _authorized_backend_json(
    "GET",
    _chat_result_url(request_id),
    timeout_seconds=timeout_seconds,
  )
  state = response.get("state")
  if state not in {"pending", "done", "failed", "timeout"}:
    raise RuntimeError("Chat relay poll response returned an invalid state")
  return response


async def _send_chat_message(
  text: str,
  *,
  force_compaction: bool = False,
  timeout_s: Any = None,
  seed_history: Any = None,
  model: str | None = None,
  approve_tool_classes: Any = None,
  approval_window_seconds: Any = None,
  workbook: Any = None,
) -> Dict[str, Any]:
  if not _chat_relay_dev_enabled():
    raise RuntimeError(f"{CHAT_RELAY_DEV_ENV}=1 is required to use {CHAT_RELAY_TOOL_NAME}")
  if not isinstance(text, str):
    raise RuntimeError("text must be a string")
  if model is not None and not isinstance(model, str):
    raise RuntimeError("model must be a string")
  if workbook is not None and not isinstance(workbook, str):
    raise RuntimeError("workbook must be a string")

  timeout_value = _coerce_chat_timeout(timeout_s)
  seed_history_value = _coerce_seed_history(seed_history)
  model_value = model.strip() if isinstance(model, str) else None
  workbook_value = workbook.strip() if isinstance(workbook, str) else None
  submitted = await asyncio.to_thread(
    _submit_chat_message,
    text,
    bool(force_compaction),
    timeout_value,
    seed_history_value,
    model_value or None,
    approve_tool_classes,
    approval_window_seconds,
    workbook_value or None,
  )
  request_id = submitted["request_id"]
  delegation_id = submitted.get("delegation_id")
  deadline = time.monotonic() + timeout_value

  while True:
    remaining = deadline - time.monotonic()
    if remaining <= 0:
      envelope: Dict[str, Any] = {
        "request_id": request_id,
        "state": "timeout",
        "error": {"message": f"Timed out waiting for chat relay result after {timeout_value:g}s"},
      }
      if delegation_id:
        envelope["delegation_id"] = delegation_id
      return envelope

    poll_timeout = min(CHAT_RELAY_HTTP_TIMEOUT_SECONDS, max(0.25, remaining))
    result = await asyncio.to_thread(_poll_chat_message, request_id, poll_timeout)
    state = result.get("state")
    if state != "pending":
      envelope = {"request_id": request_id, **result}
      if delegation_id:
        envelope["delegation_id"] = delegation_id
      return envelope

    sleep_for = min(CHAT_RELAY_POLL_INTERVAL_SECONDS, max(0.0, deadline - time.monotonic()))
    if sleep_for > 0:
      await asyncio.sleep(sleep_for)


def _error_code(message: str) -> str:
  lowered = message.lower()
  if "no active excel taskpane connection" in lowered or "no_frontend" in lowered:
    return "no_active_excel_taskpane"
  if "unknown_session" in lowered or "unknown workbook session" in lowered:
    return "unknown_workbook_session"
  if "session_disconnected" in lowered or "disconnected" in lowered:
    return "workbook_session_disconnected"
  if "neither excel_mcp_api_key nor excel_mcp_secret" in lowered or "unauthorized" in lowered:
    return "authentication_required"
  if "timeout" in lowered or "timed out" in lowered:
    return "tool_timeout"
  if "invalid json" in lowered:
    return "invalid_backend_json"
  if "must be an object" in lowered:
    return "invalid_tool_arguments"
  return "tool_execution_failed"


def _suggested_tool_calls(tool_name: str, code: str, message: str) -> list[dict[str, Any]]:
  lowered = message.lower()
  suggestions: list[dict[str, Any]] = []
  if code in {"no_active_excel_taskpane", "unknown_workbook_session", "workbook_session_disconnected"}:
    suggestions.append({"tool_name": "list_workbooks", "arguments": {}})
  if "sheet" in lowered or "worksheet" in lowered:
    suggestions.append({"tool_name": "list_sheets", "arguments": {}})
  if tool_name != "channel_status" and code == "no_active_excel_taskpane":
    suggestions.append({"tool_name": "channel_status", "arguments": {}})
  return suggestions


def _exception_envelope(tool_name: str, exc: Exception) -> dict[str, Any]:
  message = str(exc) or exc.__class__.__name__
  code = _error_code(message)
  return {
    "status": "error",
    "code": code,
    "message": message,
    "tool_name": tool_name,
    "recoverable": code
    in {
      "no_active_excel_taskpane",
      "unknown_workbook_session",
      "workbook_session_disconnected",
      "tool_timeout",
      "invalid_tool_arguments",
    },
    "suggested_tool_calls": _suggested_tool_calls(tool_name, code, message),
  }


class RegistryTool(Tool):
  """Tool implementation backed by shared registry JSON schema and backend proxy calls."""

  _tool_name: str

  def __init__(self, *, tool_name: str, **kwargs: Any) -> None:
    super().__init__(**kwargs)
    object.__setattr__(self, "_tool_name", tool_name)

  async def run(self, arguments: dict[str, Any] | None) -> ToolResult:
    try:
      tool_args: Dict[str, Any]
      if arguments is None:
        tool_args = {}
      elif isinstance(arguments, dict):
        tool_args = dict(arguments)
      else:
        raise RuntimeError("Tool arguments must be an object")

      schema_props = self.parameters.get("properties", {}) if isinstance(self.parameters, dict) else {}
      if isinstance(schema_props, dict):
        for param_name, param_spec in schema_props.items():
          if (
            isinstance(param_name, str)
            and param_name not in tool_args
            and isinstance(param_spec, dict)
            and "default" in param_spec
          ):
            tool_args[param_name] = param_spec["default"]

      result = await asyncio.to_thread(
        _call_backend,
        self._tool_name,
        tool_args,
        DEFAULT_TIMEOUT_SECONDS,
      )
    except Exception as exc:
      envelope = _exception_envelope(self._tool_name, exc)
      return ToolResult(
        structured_content=envelope,
        content=[TextContent(type="text", text=json.dumps(envelope, sort_keys=True))],
      )

    if isinstance(result, dict):
      return ToolResult(structured_content=result)

    text = result if isinstance(result, str) else json.dumps(result)
    return ToolResult(content=[TextContent(type="text", text=text)])


class ChatRelayTool(Tool):
  """Dev-only MCP tool that submits chat relay turns and polls for completion."""

  _tool_name: str

  def __init__(self, **kwargs: Any) -> None:
    super().__init__(**kwargs)
    object.__setattr__(self, "_tool_name", CHAT_RELAY_TOOL_NAME)

  async def run(self, arguments: dict[str, Any] | None) -> ToolResult:
    try:
      tool_args: Dict[str, Any]
      if arguments is None:
        tool_args = {}
      elif isinstance(arguments, dict):
        tool_args = dict(arguments)
      else:
        raise RuntimeError("Tool arguments must be an object")

      text = tool_args.get("text")
      if not isinstance(text, str):
        raise RuntimeError("text is required and must be a string")

      result = await _send_chat_message(
        text,
        force_compaction=bool(tool_args.get("force_compaction", False)),
        timeout_s=tool_args.get("timeout_s"),
        seed_history=tool_args.get("seed_history"),
        model=tool_args.get("model"),
        approve_tool_classes=tool_args.get("approve_tool_classes"),
        approval_window_seconds=tool_args.get("approval_window_seconds"),
        workbook=tool_args.get("workbook"),
      )
    except Exception as exc:
      envelope = _exception_envelope(self._tool_name, exc)
      return ToolResult(
        structured_content=envelope,
        content=[TextContent(type="text", text=json.dumps(envelope, sort_keys=True))],
      )

    return ToolResult(structured_content=result)


async def _channel_status() -> dict:
  from ._gateway_session import default_tls_verify, get_session_token, invalidate

  tls_verify = (
    _TLS_VERIFY_RAW in _TRUTHY_ENV_VALUES
    if _TLS_VERIFY_RAW
    else default_tls_verify(GATEWAY_BASE_URL)
  )

  headers: Dict[str, str] = {}
  if USER_API_KEY:
    token = get_session_token(USER_API_KEY, gateway_base_url=GATEWAY_BASE_URL, tls_verify=tls_verify)
    headers["Authorization"] = f"Bearer {token}"
  else:
    raise RuntimeError(
      "EXCEL_MCP_API_KEY is required; set it to a channel='mcp' GATEWAY_USER_KEYS key."
    )

  try:
    with httpx.Client(verify=tls_verify, timeout=10) as client:
      response = client.get(
        f"{BACKEND_BASE_URL}/api/mcp/channel-status",
        headers=headers,
      )
  except httpx.HTTPError as exc:
    raise RuntimeError(f"Backend request failed: {exc}") from exc

  if response.status_code == 401 and USER_API_KEY:
    invalidate(USER_API_KEY, GATEWAY_BASE_URL)
    token = get_session_token(USER_API_KEY, gateway_base_url=GATEWAY_BASE_URL, tls_verify=tls_verify)
    headers["Authorization"] = f"Bearer {token}"
    try:
      with httpx.Client(verify=tls_verify, timeout=10) as client:
        response = client.get(
          f"{BACKEND_BASE_URL}/api/mcp/channel-status",
          headers=headers,
        )
    except httpx.HTTPError as exc:
      raise RuntimeError(f"Backend request failed after re-auth: {exc}") from exc

  if response.status_code >= 400:
    try:
      body = response.json()
    except Exception:
      body = {"error": response.text or f"HTTP {response.status_code}"}
    error_text = body.get("error") if isinstance(body, dict) else None
    raise RuntimeError(error_text or f"Backend request failed ({response.status_code})")

  try:
    data = response.json()
  except Exception as exc:
    raise RuntimeError("Backend returned invalid JSON") from exc

  if not isinstance(data, dict):
    raise RuntimeError("Backend response must be an object")

  return data


@mcp.tool()
async def channel_status() -> dict:
  """Check relay health and connected Excel channels with a structured recovery path.

  Discovery: call this when workbook tools report a disconnected taskpane or when
  the agent needs to distinguish backend reachability from workbook routing. The
  result is a JSON object from the gateway on success. On authentication, network,
  or missing-taskpane failures, this tool returns status=error with code, message,
  recoverable, tool_name, and suggested_tool_calls so agents can decide whether to
  open Excel, refresh auth, or run list_workbooks next.
  """
  try:
    return await _channel_status()
  except Exception as exc:
    return _exception_envelope("channel_status", exc)


def _chat_relay_tool_schema() -> Dict[str, Any]:
  return {
    "type": "object",
    "properties": {
      "text": {
        "type": "string",
        "description": "User chat message to send through the Excel taskpane chat relay.",
      },
      "force_compaction": {
        "type": "boolean",
        "default": False,
        "description": "Ask the taskpane chat handler to force compaction for this turn.",
      },
      "timeout_s": {
        "type": "number",
        "default": CHAT_RELAY_DEFAULT_TIMEOUT_SECONDS,
        "minimum": 0,
        "description": "Maximum seconds to poll for the buffered chat result.",
      },
      "seed_history": {
        "type": "array",
        "description": (
          "Dev/test-only: pre-seed a synthetic user/assistant conversation for this "
          "relay turn's compaction projection. Not persisted or rendered."
        ),
        "items": {
          "type": "object",
          "properties": {
            "role": {"type": "string", "enum": ["user", "assistant"]},
            "content": {"type": "string"},
          },
          "required": ["role", "content"],
          "additionalProperties": False,
        },
      },
      "model": {
        "type": "string",
        "description": "Optional per-turn model override. Accepts bare IDs or provider:model; the gateway validates it.",
      },
      "approve_tool_classes": {
        "type": "array",
        "description": (
          "Optional dev/test delegated approval ceiling. Omit or pass an empty array to keep "
          "the default deny-with-provenance relay policy. Allowed values are read, "
          "pure_transform, artifact_write, and state_write. external_write, irreversible, "
          "and portfolio_config are never auto-approvable; tools above the ceiling escalate "
          "and the relay turn blocks until the approval is decided or expires."
        ),
        "items": {
          "type": "string",
          "enum": ["read", "pure_transform", "artifact_write", "state_write"],
        },
      },
      "approval_window_seconds": {
        "type": "integer",
        "default": 600,
        "minimum": 1,
        "description": "Validity window for a server-minted delegated approval grant.",
      },
      "workbook": {
        "type": "string",
        "description": "Optional workbook name, workbook session, or Excel gateway session for delegated approval targeting.",
      },
    },
    "required": ["text"],
  }


def _register_dev_chat_tool(mcp_instance: FastMCP) -> bool:
  if not _chat_relay_dev_enabled():
    return False
  mcp_instance.add_tool(
    ChatRelayTool(
      name=CHAT_RELAY_TOOL_NAME,
      description=(
        "Dev-only: send a chat message through the active Excel taskpane chat relay and "
        "return the buffered result state. Approval-gated tools are denied by default with "
        "relay provenance unless approve_tool_classes requests a server-minted delegated grant."
      ),
      parameters=_chat_relay_tool_schema(),
    )
  )
  return True


for spec in get_tool_specs():
  mcp.add_tool(
    RegistryTool(
      tool_name=spec["name"],
      name=spec["name"],
      description=spec["description"],
      parameters=spec["input_schema"],
    )
  )

_register_dev_chat_tool(mcp)


__all__ = [
  "mcp",
  "RegistryTool",
  "ChatRelayTool",
  "_call_backend",
  "_send_chat_message",
  "_register_dev_chat_tool",
  "_prepare_stdio_instance",
  "_kill_previous_instance",
  "channel_status",
]
