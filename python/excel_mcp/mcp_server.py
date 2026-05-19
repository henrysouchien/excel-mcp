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
MCP_SECRET = os.getenv("EXCEL_MCP_SECRET", "").strip()
DEFAULT_TIMEOUT_SECONDS = int(os.getenv("EXCEL_MCP_TOOL_TIMEOUT", "60"))
_TRUTHY_ENV_VALUES = {"1", "true", "yes", "on"}

mcp = FastMCP(
  MCP_NAME,
  instructions="Excel workbook read/write tools via the local Excel add-in taskpane.",
)


def _derive_backend_base_url(backend_url: str) -> str:
  parsed = urlparse(backend_url)
  if parsed.scheme and parsed.netloc:
    return urlunparse((parsed.scheme, parsed.netloc, "", "", "", "")).rstrip("/")
  return backend_url.strip().rstrip("/")


if not BACKEND_BASE_URL:
  BACKEND_BASE_URL = _derive_backend_base_url(BACKEND_URL)


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
  if not MCP_SECRET:
    raise RuntimeError("EXCEL_MCP_SECRET is not set")

  payload = {
    "tool_name": tool_name,
    "tool_input": tool_input,
  }
  headers = {
    "Content-Type": "application/json",
    "X-MCP-Secret": MCP_SECRET,
  }

  try:
    with httpx.Client(verify=False, timeout=timeout_seconds) as client:
      response = client.post(BACKEND_URL, json=payload, headers=headers)
  except httpx.HTTPError as exc:
    raise RuntimeError(f"Backend request failed: {exc}") from exc

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


class RegistryTool(Tool):
  """Tool implementation backed by shared registry JSON schema and backend proxy calls."""

  _tool_name: str

  def __init__(self, *, tool_name: str, **kwargs: Any) -> None:
    super().__init__(**kwargs)
    object.__setattr__(self, "_tool_name", tool_name)

  async def run(self, arguments: dict[str, Any] | None) -> ToolResult:
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

    if isinstance(result, dict):
      return ToolResult(structured_content=result)

    text = result if isinstance(result, str) else json.dumps(result)
    return ToolResult(content=[TextContent(type="text", text=text)])


async def _channel_status() -> dict:
  if not MCP_SECRET:
    raise RuntimeError("EXCEL_MCP_SECRET is not set")

  try:
    with httpx.Client(verify=False, timeout=10) as client:
      response = client.get(
        f"{BACKEND_BASE_URL}/api/mcp/channel-status",
        headers={"X-MCP-Secret": MCP_SECRET},
      )
  except httpx.HTTPError as exc:
    raise RuntimeError(f"Backend request failed: {exc}") from exc

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
  """Check which channels are connected to the relay (Excel, MCP, etc.)."""
  return await _channel_status()


for spec in get_tool_specs():
  mcp.add_tool(
    RegistryTool(
      tool_name=spec["name"],
      name=spec["name"],
      description=spec["description"],
      parameters=spec["input_schema"],
    )
  )


__all__ = [
  "mcp",
  "RegistryTool",
  "_call_backend",
  "_prepare_stdio_instance",
  "_kill_previous_instance",
  "channel_status",
]
