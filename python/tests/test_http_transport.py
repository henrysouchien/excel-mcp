from __future__ import annotations

from typing import Any, Dict

from fastapi.testclient import TestClient

from excel_mcp.relay import create_relay_app


SECRET = "test-secret"


def make_app(monkeypatch):  # type: ignore[no-untyped-def]
  monkeypatch.setenv("ENVIRONMENT", "development")
  monkeypatch.setenv("EXCEL_MCP_SECRET", SECRET)
  app = create_relay_app()
  headers = {"X-MCP-Secret": SECRET}
  return app, headers


def install_execute_spy(app):  # type: ignore[no-untyped-def]
  captured: Dict[str, Any] = {}

  async def fake_execute(
    tool_name: str,
    tool_input: Dict[str, Any],
    timeout: int,
    target_session: str | None = None,
  ) -> Dict[str, Any]:
    captured["tool_name"] = tool_name
    captured["tool_input"] = dict(tool_input)
    captured["timeout"] = timeout
    captured["target_session"] = target_session
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
