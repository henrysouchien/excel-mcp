from __future__ import annotations

import json
import threading
from typing import Any

import pytest

from excel_mcp import _gateway_session as gateway_session


class _FakeResponse:
  def __init__(self, status_code: int, payload: dict[str, Any]) -> None:
    self.status_code = status_code
    self._payload = payload
    self.text = json.dumps(payload)

  def json(self) -> dict[str, Any]:
    return dict(self._payload)


class _InitClientTransport:
  def __init__(self, responses: list[_FakeResponse]) -> None:
    self.responses = list(responses)
    self.calls: list[dict[str, Any]] = []
    self.client_kwargs: list[dict[str, Any]] = []
    self._lock = threading.Lock()

  def next_response(self) -> _FakeResponse:
    with self._lock:
      if not self.responses:
        raise AssertionError("No fake response configured")
      return self.responses.pop(0)


class _FakeInitClient:
  def __init__(self, transport: _InitClientTransport, kwargs: dict[str, Any]) -> None:
    self._transport = transport
    self._transport.client_kwargs.append(kwargs)

  def __enter__(self):
    return self

  def __exit__(self, exc_type, exc, tb):
    return False

  def post(self, url: str, json: dict[str, Any]) -> _FakeResponse:
    self._transport.calls.append({"url": url, "json": dict(json)})
    return self._transport.next_response()


def _response(token: str, *, expires_at: int = 10_000) -> _FakeResponse:
  return _FakeResponse(
    200,
    {
      "session_token": token,
      "session_id": f"session-{token}",
      "expires_at": expires_at,
    },
  )


def _install_client(monkeypatch, responses: list[_FakeResponse]) -> _InitClientTransport:
  transport = _InitClientTransport(responses)

  def _client_factory(**kwargs: Any) -> _FakeInitClient:
    return _FakeInitClient(transport, kwargs)

  monkeypatch.setattr(gateway_session.httpx, "Client", _client_factory)
  return transport


@pytest.fixture(autouse=True)
def _clear_cache() -> None:
  with gateway_session._LOCK:
    gateway_session._CACHE.clear()


def test_get_session_token_initial_exchange(monkeypatch) -> None:
  monkeypatch.setattr(gateway_session.time, "time", lambda: 100)
  transport = _install_client(monkeypatch, [_response("jwt-1")])

  token = gateway_session.get_session_token(
    "api-key-1",
    gateway_base_url="https://localhost:8000",
    tls_verify=False,
  )

  assert token == "jwt-1"
  assert len(transport.calls) == 1
  assert transport.calls[0]["url"] == "https://localhost:8000/api/chat/init"
  assert transport.calls[0]["json"] == {
    "api_key": "api-key-1",
    "user_id": "mcp-subprocess",
    "context": {"channel": "mcp"},
  }
  assert len(gateway_session._CACHE) == 1


def test_get_session_token_uses_cache_when_fresh(monkeypatch) -> None:
  monkeypatch.setattr(gateway_session.time, "time", lambda: 100)
  transport = _install_client(monkeypatch, [_response("jwt-1")])

  first = gateway_session.get_session_token("api-key-1", gateway_base_url="https://localhost:8000")
  second = gateway_session.get_session_token("api-key-1", gateway_base_url="https://localhost:8000")

  assert first == "jwt-1"
  assert second == "jwt-1"
  assert len(transport.calls) == 1


def test_get_session_token_refreshes_near_expiry(monkeypatch) -> None:
  monkeypatch.setattr(gateway_session.time, "time", lambda: 1_000)
  transport = _install_client(
    monkeypatch,
    [
      _response("jwt-1", expires_at=1_030),
      _response("jwt-2", expires_at=2_000),
    ],
  )

  first = gateway_session.get_session_token("api-key-1", gateway_base_url="https://localhost:8000")
  second = gateway_session.get_session_token("api-key-1", gateway_base_url="https://localhost:8000")

  assert first == "jwt-1"
  assert second == "jwt-2"
  assert len(transport.calls) == 2


def test_get_session_token_separate_keys_cached_independently(monkeypatch) -> None:
  monkeypatch.setattr(gateway_session.time, "time", lambda: 100)
  transport = _install_client(monkeypatch, [_response("jwt-1"), _response("jwt-2")])

  first = gateway_session.get_session_token("api-key-1", gateway_base_url="https://localhost:8000")
  second = gateway_session.get_session_token("api-key-2", gateway_base_url="https://localhost:8000")

  assert first == "jwt-1"
  assert second == "jwt-2"
  assert len(transport.calls) == 2
  assert len(gateway_session._CACHE) == 2


def test_invalidate_drops_cache_entry(monkeypatch) -> None:
  monkeypatch.setattr(gateway_session.time, "time", lambda: 100)
  transport = _install_client(monkeypatch, [_response("jwt-1"), _response("jwt-2")])

  first = gateway_session.get_session_token("api-key-1", gateway_base_url="https://localhost:8000")
  gateway_session.invalidate("api-key-1", "https://localhost:8000")
  second = gateway_session.get_session_token("api-key-1", gateway_base_url="https://localhost:8000")

  assert first == "jwt-1"
  assert second == "jwt-2"
  assert len(transport.calls) == 2


def test_exchange_raises_on_4xx(monkeypatch) -> None:
  monkeypatch.setattr(gateway_session.time, "time", lambda: 100)
  _install_client(monkeypatch, [_FakeResponse(401, {"error": "secret echo should not appear"})])

  with pytest.raises(RuntimeError) as exc:
    gateway_session.get_session_token("api-key-1", gateway_base_url="https://localhost:8000")

  assert str(exc.value) == "Gateway /api/chat/init failed: HTTP 401"
  assert "secret echo" not in str(exc.value)
  assert gateway_session._CACHE == {}


def test_concurrent_calls_under_lock(monkeypatch) -> None:
  monkeypatch.setattr(gateway_session.time, "time", lambda: 100)
  transport = _install_client(monkeypatch, [_response("jwt-1")])
  start = threading.Event()
  all_ready = threading.Event()
  ready_lock = threading.Lock()
  ready_count = 0
  results: list[str] = []
  errors: list[BaseException] = []

  def _worker() -> None:
    nonlocal ready_count
    with ready_lock:
      ready_count += 1
      if ready_count == 8:
        all_ready.set()
    start.wait()
    try:
      results.append(gateway_session.get_session_token("api-key-1", gateway_base_url="https://localhost:8000"))
    except BaseException as exc:
      errors.append(exc)

  threads = [threading.Thread(target=_worker) for _ in range(8)]
  for thread in threads:
    thread.start()

  assert all_ready.wait(timeout=2)
  start.set()
  for thread in threads:
    thread.join(timeout=2)

  assert errors == []
  assert results == ["jwt-1"] * 8
  assert len(transport.calls) == 1
