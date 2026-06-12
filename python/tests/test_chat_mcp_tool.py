from __future__ import annotations

import asyncio
from typing import Any, Dict

import pytest

from excel_mcp import mcp_server
from excel_mcp.tool_registry import get_tool_names


def run_async(coro):  # type: ignore[no-untyped-def]
  return asyncio.run(coro)


def test_dev_chat_tool_registration_is_gated(monkeypatch: pytest.MonkeyPatch) -> None:
  class _FakeMcp:
    def __init__(self) -> None:
      self.tools: list[Any] = []

    def add_tool(self, tool: Any) -> None:
      self.tools.append(tool)

  monkeypatch.delenv("EXCEL_CHAT_RELAY_DEV", raising=False)
  disabled = _FakeMcp()
  assert mcp_server._register_dev_chat_tool(disabled) is False
  assert disabled.tools == []

  monkeypatch.setenv("EXCEL_CHAT_RELAY_DEV", "1")
  enabled = _FakeMcp()
  assert mcp_server._register_dev_chat_tool(enabled) is True
  assert len(enabled.tools) == 1
  assert getattr(enabled.tools[0], "_tool_name") == "send_chat_message"


def test_send_chat_message_is_not_in_shared_excel_tool_registry() -> None:
  assert "send_chat_message" not in get_tool_names()


def test_send_chat_message_requires_dev_flag(monkeypatch: pytest.MonkeyPatch) -> None:
  monkeypatch.delenv("EXCEL_CHAT_RELAY_DEV", raising=False)

  with pytest.raises(RuntimeError, match="EXCEL_CHAT_RELAY_DEV=1"):
    run_async(mcp_server._send_chat_message("hello", timeout_s=0.1))


def test_send_chat_message_submits_then_polls_until_done(monkeypatch: pytest.MonkeyPatch) -> None:
  monkeypatch.setenv("EXCEL_CHAT_RELAY_DEV", "1")
  monkeypatch.setattr(mcp_server, "CHAT_RELAY_POLL_INTERVAL_SECONDS", 0)
  captured: Dict[str, Any] = {}
  poll_responses = iter([
    {"state": "pending"},
    {"state": "done", "result": {"assistant_text": "hi"}},
  ])

  seed_history = [
    {"role": "user", "content": "first"},
    {"role": "assistant", "content": "second"},
  ]

  def _submit(
    text: str,
    force_compaction: bool,
    timeout_s: float,
    seed_history: list[Dict[str, str]] | None = None,
    model: str | None = None,
    approve_tool_classes: list[str] | None = None,
    approval_window_seconds: Any = None,
    workbook: str | None = None,
  ) -> Dict[str, str]:
    captured["submit"] = {
      "text": text,
      "force_compaction": force_compaction,
      "timeout_s": timeout_s,
      "seed_history": seed_history,
      "model": model,
      "approve_tool_classes": approve_tool_classes,
      "approval_window_seconds": approval_window_seconds,
      "workbook": workbook,
    }
    return {"request_id": "req-chat-1"}

  def _poll(request_id: str, timeout_seconds: float = 10) -> Dict[str, Any]:
    captured.setdefault("polls", []).append({
      "request_id": request_id,
      "timeout_seconds": timeout_seconds,
    })
    return next(poll_responses)

  monkeypatch.setattr(mcp_server, "_submit_chat_message", _submit)
  monkeypatch.setattr(mcp_server, "_poll_chat_message", _poll)

  result = run_async(
    mcp_server._send_chat_message(
      "hello",
      force_compaction=True,
      timeout_s=5,
      seed_history=seed_history,
      model="claude-sonnet-4-6",
    )
  )

  assert result == {
    "request_id": "req-chat-1",
    "state": "done",
    "result": {"assistant_text": "hi"},
  }
  assert captured["submit"] == {
    "text": "hello",
    "force_compaction": True,
    "timeout_s": 5.0,
    "seed_history": seed_history,
    "model": "claude-sonnet-4-6",
    "approve_tool_classes": None,
    "approval_window_seconds": None,
    "workbook": None,
  }
  assert [item["request_id"] for item in captured["polls"]] == ["req-chat-1", "req-chat-1"]


def test_send_chat_message_delegated_submit_returns_delegation_id_when_present(monkeypatch: pytest.MonkeyPatch) -> None:
  monkeypatch.setenv("EXCEL_CHAT_RELAY_DEV", "1")
  monkeypatch.setattr(mcp_server, "CHAT_RELAY_POLL_INTERVAL_SECONDS", 0)
  captured: Dict[str, Any] = {}

  def _submit(
    text: str,
    force_compaction: bool,
    timeout_s: float,
    seed_history: list[Dict[str, str]] | None = None,
    model: str | None = None,
    approve_tool_classes: list[str] | None = None,
    approval_window_seconds: Any = None,
    workbook: str | None = None,
  ) -> Dict[str, str]:
    captured["submit"] = {
      "text": text,
      "force_compaction": force_compaction,
      "timeout_s": timeout_s,
      "seed_history": seed_history,
      "model": model,
      "approve_tool_classes": approve_tool_classes,
      "approval_window_seconds": approval_window_seconds,
      "workbook": workbook,
    }
    return {"request_id": "req-chat-2", "delegation_id": "delegation-2"}

  def _poll(request_id: str, timeout_seconds: float = 10) -> Dict[str, Any]:
    captured["poll"] = {"request_id": request_id, "timeout_seconds": timeout_seconds}
    return {"state": "done", "result": {"assistant_text": "done"}}

  monkeypatch.setattr(mcp_server, "_submit_chat_message", _submit)
  monkeypatch.setattr(mcp_server, "_poll_chat_message", _poll)

  result = run_async(
    mcp_server._send_chat_message(
      "delegate",
      timeout_s=5,
      approve_tool_classes=["read", "state_write"],
      approval_window_seconds=120,
      workbook="Budget.xlsx",
    )
  )

  assert result == {
    "request_id": "req-chat-2",
    "state": "done",
    "result": {"assistant_text": "done"},
    "delegation_id": "delegation-2",
  }
  assert captured["submit"] == {
    "text": "delegate",
    "force_compaction": False,
    "timeout_s": 5.0,
    "seed_history": [],
    "model": None,
    "approve_tool_classes": ["read", "state_write"],
    "approval_window_seconds": 120,
    "workbook": "Budget.xlsx",
  }
  assert captured["poll"]["request_id"] == "req-chat-2"


def test_send_chat_message_rejects_invalid_seed_history(monkeypatch: pytest.MonkeyPatch) -> None:
  monkeypatch.setenv("EXCEL_CHAT_RELAY_DEV", "1")

  with pytest.raises(RuntimeError, match=r"seed_history\[0\]\.role"):
    run_async(
      mcp_server._send_chat_message(
        "hello",
        timeout_s=0.1,
        seed_history=[{"role": "system", "content": "nope"}],
      )
    )


def test_send_chat_message_schema_documents_seed_history_and_model() -> None:
  properties = mcp_server._chat_relay_tool_schema()["properties"]
  seed_history = properties["seed_history"]

  assert seed_history["type"] == "array"
  assert "Dev/test-only" in seed_history["description"]
  assert seed_history["items"]["properties"]["role"]["enum"] == ["user", "assistant"]
  assert properties["model"]["type"] == "string"
  assert "provider:model" in properties["model"]["description"]
  assert properties["approve_tool_classes"]["type"] == "array"
  assert properties["approve_tool_classes"]["items"]["enum"] == ["read", "pure_transform", "artifact_write", "state_write"]
  assert "default deny-with-provenance" in properties["approve_tool_classes"]["description"]
  assert "never auto-approvable" in properties["approve_tool_classes"]["description"]
  assert properties["approval_window_seconds"]["default"] == 600
  assert properties["approval_window_seconds"]["minimum"] == 1
  assert properties["workbook"]["type"] == "string"
