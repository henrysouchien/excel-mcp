from __future__ import annotations

from typing import Any, Dict

from fastapi.testclient import TestClient

from excel_mcp.relay import RelayAuthContext, create_relay_app


def make_app(monkeypatch):  # type: ignore[no-untyped-def]
  del monkeypatch

  async def authenticate_request(_request):
    return RelayAuthContext(session_id="gateway-session-1", user_id="user-1", channel="mcp")

  app = create_relay_app(authenticate_request=authenticate_request)
  headers = {"Authorization": "Bearer test-token"}
  return app, headers


def install_execute_spy(app):  # type: ignore[no-untyped-def]
  captured: Dict[str, Any] = {}

  async def fake_execute(
    tool_name: str,
    tool_input: Dict[str, Any],
    timeout: int,
    target_session: str | None = None,
    *,
    gateway_session_id: str | None = None,
    user_id: str | None = None,
    kind: str = "tool",
  ) -> Dict[str, Any]:
    captured["tool_name"] = tool_name
    captured["tool_input"] = dict(tool_input)
    captured["timeout"] = timeout
    captured["target_session"] = target_session
    captured["gateway_session_id"] = gateway_session_id
    captured["user_id"] = user_id
    captured["kind"] = kind
    return {"request_id": "req-1", "result": {"ok": True}, "error": None}

  app.state.mcp_relay_state.relay.execute = fake_execute  # type: ignore[method-assign]
  return captured


def test_execute_request_preserves_workbook_override_inside_tool_input(monkeypatch) -> None:
  app, headers = make_app(monkeypatch)
  captured = install_execute_spy(app)

  with TestClient(app) as client:
    response = client.post(
      "/api/mcp/execute",
      json={
        "tool_name": "read_cells",
        "tool_input": {"range": "A1:B2", "_workbook": "session-a"},
      },
      headers=headers,
    )

  assert response.status_code == 200
  assert response.json()["result"] == {"ok": True}
  assert captured == {
    "tool_name": "read_cells",
    "tool_input": {"range": "A1:B2", "_workbook": "session-a"},
    "timeout": 60,
    "target_session": None,
    "gateway_session_id": "gateway-session-1",
    "user_id": "user-1",
    "kind": "tool",
  }


def test_execute_request_handles_tool_input_with_only_workbook_override(monkeypatch) -> None:
  app, headers = make_app(monkeypatch)
  captured = install_execute_spy(app)

  with TestClient(app) as client:
    response = client.post(
      "/api/mcp/execute",
      json={
        "tool_name": "list_sheets",
        "tool_input": {"_workbook": "session-a"},
      },
      headers=headers,
    )

  assert response.status_code == 200
  assert response.json()["result"] == {"ok": True}
  assert captured["tool_name"] == "list_sheets"
  assert captured["tool_input"] == {"_workbook": "session-a"}


def test_execute_request_without_workbook_override_uses_plain_tool_input_shape(monkeypatch) -> None:
  app, headers = make_app(monkeypatch)
  captured = install_execute_spy(app)

  with TestClient(app) as client:
    response = client.post(
      "/api/mcp/execute",
      json={
        "tool_name": "read_cells",
        "tool_input": {"range": "C3"},
      },
      headers=headers,
    )

  assert response.status_code == 200
  assert response.json()["result"] == {"ok": True}
  assert captured["tool_name"] == "read_cells"
  assert captured["tool_input"] == {"range": "C3"}
  assert captured["kind"] == "tool"


def test_chat_execute_rejects_when_dev_flag_unset(monkeypatch) -> None:
  monkeypatch.delenv("EXCEL_CHAT_RELAY_DEV", raising=False)
  app, headers = make_app(monkeypatch)

  with TestClient(app) as client:
    response = client.post(
      "/api/mcp/execute",
      json={
        "kind": "chat",
        "tool_name": "send_chat_message",
        "tool_input": {"text": "hello", "force_compaction": False},
      },
      headers=headers,
    )

  assert response.status_code == 403
  assert response.json() == {"error": "Chat relay is disabled"}


def test_chat_execute_submits_and_result_polls_when_dev_flag_set(monkeypatch) -> None:
  monkeypatch.setenv("EXCEL_CHAT_RELAY_DEV", "1")
  app, headers = make_app(monkeypatch)
  captured: Dict[str, Any] = {}

  class _FakeRelay:
    async def submit(
      self,
      tool_name: str,
      tool_input: Dict[str, Any],
      timeout: int,
      target_session: str | None = None,
      *,
      gateway_session_id: str | None = None,
      user_id: str | None = None,
      kind: str = "tool",
      request_id: str | None = None,
    ) -> Dict[str, Any]:
      captured["submit"] = {
        "tool_name": tool_name,
        "tool_input": dict(tool_input),
        "timeout": timeout,
        "target_session": target_session,
        "gateway_session_id": gateway_session_id,
        "user_id": user_id,
        "kind": kind,
      }
      return {"request_id": "req-chat-1"}

    async def result(
      self,
      request_id: str,
      *,
      gateway_session_id: str | None = None,
      user_id: str | None = None,
    ):
      captured["result"] = {
        "request_id": request_id,
        "gateway_session_id": gateway_session_id,
        "user_id": user_id,
      }
      return "ok", {"state": "pending"}

  app.state.mcp_relay_state.relay = _FakeRelay()

  with TestClient(app) as client:
    submit_response = client.post(
      "/api/mcp/execute",
      json={
        "kind": "chat",
        "tool_name": "send_chat_message",
        "tool_input": {"text": "hello", "force_compaction": True, "timeout_s": 12.2},
      },
      headers=headers,
    )
    poll_response = client.get("/api/mcp/result/req-chat-1", headers=headers)

  assert submit_response.status_code == 200
  assert submit_response.json() == {"request_id": "req-chat-1"}
  assert poll_response.status_code == 200
  assert poll_response.json() == {"state": "pending"}
  assert captured["submit"] == {
    "tool_name": "send_chat_message",
    "tool_input": {"text": "hello", "force_compaction": True, "timeout_s": 12.2},
    "timeout": 13,
    "target_session": None,
    "gateway_session_id": "gateway-session-1",
    "user_id": "user-1",
    "kind": "chat",
  }
  assert captured["result"] == {
    "request_id": "req-chat-1",
    "gateway_session_id": "gateway-session-1",
    "user_id": "user-1",
  }
