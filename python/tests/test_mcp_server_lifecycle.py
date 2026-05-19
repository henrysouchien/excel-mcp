from __future__ import annotations

from pathlib import Path

from excel_mcp import mcp_server


def _use_pid_file(monkeypatch, pid_file: Path) -> None:
  monkeypatch.setenv("EXCEL_MCP_PID_FILE", str(pid_file))
  monkeypatch.setattr(mcp_server.os, "getpid", lambda: 22222)


def test_stdio_prepare_does_not_kill_existing_pid_by_default(tmp_path, monkeypatch) -> None:
  pid_file = tmp_path / "excel-mcp.pid"
  pid_file.write_text("11111")
  kills: list[tuple[int, int]] = []
  _use_pid_file(monkeypatch, pid_file)
  monkeypatch.delenv("EXCEL_MCP_SINGLETON", raising=False)
  monkeypatch.setattr(mcp_server.os, "kill", lambda pid, sig: kills.append((pid, sig)))

  mcp_server._prepare_stdio_instance()

  assert kills == []
  assert pid_file.read_text() == "22222"


def test_stdio_prepare_kills_existing_pid_when_singleton_enabled(tmp_path, monkeypatch) -> None:
  pid_file = tmp_path / "excel-mcp.pid"
  pid_file.write_text("11111")
  kills: list[tuple[int, int]] = []
  _use_pid_file(monkeypatch, pid_file)
  monkeypatch.setenv("EXCEL_MCP_SINGLETON", "1")
  monkeypatch.setattr(mcp_server.os, "kill", lambda pid, sig: kills.append((pid, sig)))

  mcp_server._prepare_stdio_instance()

  assert kills == [(11111, mcp_server.signal.SIGTERM)]
  assert pid_file.read_text() == "22222"


def test_explicit_kill_previous_instance_preserves_legacy_behavior(tmp_path, monkeypatch) -> None:
  pid_file = tmp_path / "excel-mcp.pid"
  pid_file.write_text("11111")
  kills: list[tuple[int, int]] = []
  _use_pid_file(monkeypatch, pid_file)
  monkeypatch.delenv("EXCEL_MCP_SINGLETON", raising=False)
  monkeypatch.setattr(mcp_server.os, "kill", lambda pid, sig: kills.append((pid, sig)))

  mcp_server._kill_previous_instance()

  assert kills == [(11111, mcp_server.signal.SIGTERM)]
  assert pid_file.read_text() == "22222"
