from __future__ import annotations

import json
from typing import Any

import pytest

from excel_mcp import _gateway_session as gateway_session
from excel_mcp import mcp_server


class _FakeBackendResponse:
  def __init__(self, status_code: int, payload: dict[str, Any]) -> None:
    self.status_code = status_code
    self._payload = payload
    self.text = json.dumps(payload)

  def json(self) -> dict[str, Any]:
    return dict(self._payload)


class _BackendTransport:
  def __init__(self, responses: list[_FakeBackendResponse]) -> None:
    self.responses = list(responses)
    self.requests: list[dict[str, Any]] = []
    self.client_kwargs: list[dict[str, Any]] = []

  def next_response(self) -> _FakeBackendResponse:
    if not self.responses:
      raise AssertionError("No fake backend response configured")
    return self.responses.pop(0)


class _FakeBackendClient:
  def __init__(self, transport: _BackendTransport, kwargs: dict[str, Any]) -> None:
    self._transport = transport
    self._transport.client_kwargs.append(kwargs)

  def __enter__(self):
    return self

  def __exit__(self, exc_type, exc, tb):
    return False

  def post(self, url: str, json: dict[str, Any], headers: dict[str, Any]) -> _FakeBackendResponse:
    self._transport.requests.append(
      {
        "url": url,
        "json": dict(json),
        "headers": dict(headers),
      }
    )
    return self._transport.next_response()


def _install_backend_client(monkeypatch, responses: list[_FakeBackendResponse]) -> _BackendTransport:
  transport = _BackendTransport(responses)

  def _client_factory(**kwargs: Any) -> _FakeBackendClient:
    return _FakeBackendClient(transport, kwargs)

  monkeypatch.setattr(mcp_server.httpx, "Client", _client_factory)
  return transport


@pytest.fixture(autouse=True)
def _reset_auth(monkeypatch) -> None:
  monkeypatch.setattr(mcp_server, "BACKEND_URL", "https://localhost:8000/api/mcp/execute")
  monkeypatch.setattr(mcp_server, "BACKEND_BASE_URL", "https://localhost:8000")
  monkeypatch.setattr(mcp_server, "GATEWAY_BASE_URL", "https://localhost:8000")
  monkeypatch.setattr(mcp_server, "_TLS_VERIFY_RAW", "")
  monkeypatch.setattr(mcp_server, "USER_API_KEY", "")
  with gateway_session._LOCK:
    gateway_session._CACHE.clear()


def test_call_backend_uses_jwt_when_api_key_set(monkeypatch) -> None:
  monkeypatch.setattr(mcp_server, "USER_API_KEY", "user-api-key")
  monkeypatch.setattr(
    gateway_session,
    "get_session_token",
    lambda api_key, *, gateway_base_url, tls_verify: "jwt-token",
  )
  transport = _install_backend_client(monkeypatch, [_FakeBackendResponse(200, {"result": {"ok": True}})])

  result = mcp_server._call_backend("read_cells", {"range": "A1"}, timeout_seconds=10)

  assert result == {"ok": True}
  assert transport.requests[0]["headers"]["Authorization"] == "Bearer jwt-token"
  assert "X-MCP-Secret" not in transport.requests[0]["headers"]


def test_call_backend_raises_when_api_key_unset() -> None:
  with pytest.raises(RuntimeError) as exc:
    mcp_server._call_backend("read_cells", {"range": "A1"}, timeout_seconds=10)

  assert "EXCEL_MCP_API_KEY is required" in str(exc.value)


def test_call_backend_retries_once_on_401_with_jwt(monkeypatch) -> None:
  monkeypatch.setattr(mcp_server, "USER_API_KEY", "user-api-key")
  token_calls: list[dict[str, Any]] = []
  invalidations: list[tuple[str, str]] = []
  tokens = iter(["jwt-token-1", "jwt-token-2"])

  def _get_session_token(api_key: str, *, gateway_base_url: str, tls_verify: bool) -> str:
    token_calls.append(
      {
        "api_key": api_key,
        "gateway_base_url": gateway_base_url,
        "tls_verify": tls_verify,
      }
    )
    return next(tokens)

  def _invalidate(api_key: str, gateway_base_url: str) -> None:
    invalidations.append((api_key, gateway_base_url))

  monkeypatch.setattr(gateway_session, "get_session_token", _get_session_token)
  monkeypatch.setattr(gateway_session, "invalidate", _invalidate)
  transport = _install_backend_client(
    monkeypatch,
    [
      _FakeBackendResponse(401, {"error": "expired"}),
      _FakeBackendResponse(200, {"result": {"ok": True}}),
    ],
  )

  result = mcp_server._call_backend("read_cells", {"range": "A1"}, timeout_seconds=10)

  assert result == {"ok": True}
  assert invalidations == [("user-api-key", "https://localhost:8000")]
  assert len(token_calls) == 2
  assert len(transport.requests) == 2
  assert transport.requests[0]["headers"]["Authorization"] == "Bearer jwt-token-1"
  assert transport.requests[1]["headers"]["Authorization"] == "Bearer jwt-token-2"
