from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import pytest

from excel_mcp import cli


def test_ensure_package_api_keys_requires_both_channel_keys(monkeypatch, tmp_path: Path) -> None:
  monkeypatch.setattr(cli, "ENV_PATH", tmp_path / ".env")
  monkeypatch.delenv("EXCEL_MCP_API_KEY", raising=False)
  monkeypatch.delenv("SESSION_API_KEY", raising=False)

  with pytest.raises(cli.CliError, match="EXCEL_MCP_API_KEY must be exported"):
    cli._ensure_package_api_keys(force=False)

  monkeypatch.setenv("EXCEL_MCP_API_KEY", "mcp-key")
  with pytest.raises(cli.CliError, match="SESSION_API_KEY must be exported"):
    cli._ensure_package_api_keys(force=False)


def test_ensure_package_api_keys_writes_both_keys(monkeypatch, tmp_path: Path) -> None:
  env_path = tmp_path / ".env"
  monkeypatch.setattr(cli, "ENV_PATH", env_path)
  monkeypatch.setenv("EXCEL_MCP_API_KEY", "mcp-key")
  monkeypatch.setenv("SESSION_API_KEY", "excel-key")

  assert cli._ensure_package_api_keys(force=False) == "mcp-key"
  assert env_path.read_text() == "EXCEL_MCP_API_KEY=mcp-key\nSESSION_API_KEY=excel-key\n"


def test_load_package_env_requires_session_api_key(monkeypatch, tmp_path: Path) -> None:
  env_path = tmp_path / ".env"
  env_path.write_text("EXCEL_MCP_API_KEY=mcp-key\n")
  monkeypatch.setattr(cli, "ENV_PATH", env_path)

  with pytest.raises(cli.CliError, match="SESSION_API_KEY"):
    cli._load_package_env()


def test_command_start_launches_authenticated_product_gateway(monkeypatch) -> None:
  calls: list[dict[str, Any]] = []

  class _FakeProcess:
    def poll(self) -> None:
      return None

  def _start_process(label: str, command: list[str], *, cwd: Path, env: dict[str, str]) -> _FakeProcess:
    calls.append({"label": label, "command": command, "cwd": cwd, "env": env})
    return _FakeProcess()

  def _wait_for_readiness(*, stop_event: Any, **_: Any) -> None:
    stop_event.set()

  monkeypatch.setattr(cli, "_load_package_env", lambda: {"EXCEL_MCP_API_KEY": "mcp-key", "SESSION_API_KEY": "excel-key"})
  monkeypatch.setattr(cli, "_ensure_certificates", lambda: None)
  monkeypatch.setattr(cli, "_ensure_port_available", lambda _port: None)
  monkeypatch.setattr(cli, "_start_process", _start_process)
  monkeypatch.setattr(cli, "_wait_for_readiness", _wait_for_readiness)
  monkeypatch.setattr(cli, "_stop_processes", lambda *_args, **_kwargs: None)

  assert cli.command_start(argparse.Namespace()) == 0
  assert calls[0]["label"] == "gateway"
  assert calls[0]["command"][4] == "api.main:app"
  assert calls[0]["cwd"] == cli.REPO_ROOT
  assert calls[1]["label"] == "addin"
  assert cli.ADDIN_PORT == 3102
