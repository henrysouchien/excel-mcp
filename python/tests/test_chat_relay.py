from __future__ import annotations

import asyncio
from typing import Any, Dict

from excel_mcp.relay import McpExecuteRequest, McpRelay


def run_async(coro):  # type: ignore[no-untyped-def]
  return asyncio.run(coro)


async def _register(relay: McpRelay):
  return await relay.register_client("session-a", "A.xlsx", user_id="alice")


async def _event(queue: "asyncio.Queue[Dict[str, Any]]") -> Dict[str, Any]:
  return await asyncio.wait_for(queue.get(), timeout=0.25)


def test_execute_request_kind_defaults_to_tool() -> None:
  request = McpExecuteRequest(tool_name="read_cells", tool_input={"range": "A1"})

  assert request.kind == "tool"


def test_chat_kind_emits_chat_request_event() -> None:
  async def scenario() -> None:
    relay = McpRelay()
    client = await _register(relay)
    seed_history = [
      {"role": "user", "content": "seed user"},
      {"role": "assistant", "content": "seed assistant"},
    ]

    submitted = await relay.submit(
      "send_chat_message",
      {
        "text": "hello",
        "force_compaction": True,
        "seed_history": seed_history,
        "model": "claude-sonnet-4-6",
      },
      timeout=1,
      gateway_session_id="session-a",
      user_id="alice",
      kind="chat",
    )
    event = await _event(client.queue)

    assert submitted["request_id"] == event["request_id"]
    assert event["type"] == "mcp_chat_request"
    assert event["text"] == "hello"
    assert event["force_compaction"] is True
    assert event["seed_history"] == seed_history
    assert event["model"] == "claude-sonnet-4-6"
    assert event["nonce"]
    assert event["delivery_id"] == client.client_id
    assert event["replay"] is False
    assert event["delegation_id"] is None
    assert "tool_name" not in event
    assert "tool_input" not in event

  run_async(scenario())


def test_chat_kind_accepts_caller_supplied_request_id() -> None:
  async def scenario() -> None:
    relay = McpRelay()
    client = await _register(relay)

    submitted = await relay.submit(
      "send_chat_message",
      {"text": "reserved"},
      timeout=1,
      gateway_session_id="session-a",
      user_id="alice",
      kind="chat",
      request_id="relay-request-1",
    )
    event = await _event(client.queue)

    assert submitted["request_id"] == "relay-request-1"
    assert event["request_id"] == "relay-request-1"
    assert "relay-request-1" in relay._inflight
    assert "seed_history" not in event
    assert "model" not in event

  run_async(scenario())


def test_chat_execute_request_model_threads_request_id_for_route_copies() -> None:
  async def scenario() -> None:
    relay = McpRelay()
    client = await _register(relay)
    request = McpExecuteRequest(
      tool_name="send_chat_message",
      tool_input={"text": "model reserved"},
      kind="chat",
      request_id="model-route-request",
    )

    submitted = await relay.submit(
      request.tool_name,
      request.tool_input,
      timeout=1,
      gateway_session_id="session-a",
      user_id="alice",
      kind=request.kind,
    )
    event = await _event(client.queue)

    assert submitted["request_id"] == "model-route-request"
    assert event["request_id"] == "model-route-request"
    assert event["text"] == "model reserved"

  run_async(scenario())


def test_chat_kind_rejects_duplicate_caller_supplied_request_id_without_overwrite() -> None:
  async def scenario() -> None:
    relay = McpRelay()
    client = await _register(relay)

    first = await relay.submit(
      "send_chat_message",
      {"text": "first"},
      timeout=1,
      gateway_session_id="session-a",
      user_id="alice",
      kind="chat",
      request_id="duplicate-request",
    )
    await _event(client.queue)

    try:
      await relay.submit(
        "send_chat_message",
        {"text": "second"},
        timeout=1,
        gateway_session_id="session-a",
        user_id="alice",
        kind="chat",
        request_id="duplicate-request",
      )
    except RuntimeError as exc:
      assert str(exc) == "request_id_inflight"
    else:
      raise AssertionError("duplicate request_id was accepted")

    assert first["request_id"] == "duplicate-request"
    assert relay._inflight["duplicate-request"]["tool_input"]["text"] == "first"

  run_async(scenario())


def test_chat_kind_delivers_delegation_id_when_supplied() -> None:
  async def scenario() -> None:
    relay = McpRelay()
    client = await _register(relay)

    submitted = await relay.submit(
      "send_chat_message",
      {"text": "delegate", "delegation_id": "delegation-1"},
      timeout=1,
      gateway_session_id="session-a",
      user_id="alice",
      kind="chat",
    )
    event = await _event(client.queue)

    assert event["request_id"] == submitted["request_id"]
    assert event["type"] == "mcp_chat_request"
    assert event["delegation_id"] == "delegation-1"

  run_async(scenario())


def test_tool_kind_delivery_shape_is_unchanged_by_delegation_id_input() -> None:
  async def scenario() -> None:
    relay = McpRelay()
    client = await _register(relay)

    submitted = await relay.submit(
      "read_cells",
      {"range": "A1", "delegation_id": "not-chat-delegation"},
      timeout=1,
      gateway_session_id="session-a",
      user_id="alice",
      kind="tool",
    )
    event = await _event(client.queue)

    assert event["request_id"] == submitted["request_id"]
    assert event["type"] == "mcp_tool_request"
    assert event["tool_name"] == "read_cells"
    assert event["tool_input"] == {"range": "A1", "delegation_id": "not-chat-delegation"}
    assert "delegation_id" not in event

  run_async(scenario())


def test_async_result_states_pending_done_failed_timeout() -> None:
  async def scenario() -> None:
    relay = McpRelay()
    client = await _register(relay)

    pending_submit = await relay.submit(
      "send_chat_message",
      {"text": "pending", "force_compaction": False},
      timeout=1,
      gateway_session_id="session-a",
      user_id="alice",
      kind="chat",
    )
    pending_event = await _event(client.queue)
    status, pending = await relay.result(
      pending_submit["request_id"],
      gateway_session_id="session-a",
      user_id="alice",
    )
    assert status == "ok"
    assert pending == {"state": "pending"}

    assert await relay.complete(
      pending_event["request_id"],
      pending_event["nonce"],
      pending_event["delivery_id"],
      {"assistant_text": "done"},
      None,
      gateway_session_id="session-a",
      user_id="alice",
    ) == "ok"
    status, done = await relay.result(
      pending_submit["request_id"],
      gateway_session_id="session-a",
      user_id="alice",
    )
    assert status == "ok"
    assert done == {"state": "done", "result": {"assistant_text": "done"}}

    failed_submit = await relay.submit(
      "send_chat_message",
      {"text": "failed", "force_compaction": False},
      timeout=1,
      gateway_session_id="session-a",
      user_id="alice",
      kind="chat",
    )
    failed_event = await _event(client.queue)
    assert await relay.complete(
      failed_event["request_id"],
      failed_event["nonce"],
      failed_event["delivery_id"],
      None,
      {"message": "boom"},
      gateway_session_id="session-a",
      user_id="alice",
    ) == "ok"
    status, failed = await relay.result(
      failed_submit["request_id"],
      gateway_session_id="session-a",
      user_id="alice",
    )
    assert status == "ok"
    assert failed == {"state": "failed", "error": {"message": "boom"}}

    timeout_submit = await relay.submit(
      "send_chat_message",
      {"text": "timeout", "force_compaction": False},
      timeout=0.01,
      gateway_session_id="session-a",
      user_id="alice",
      kind="chat",
    )
    await _event(client.queue)
    await asyncio.sleep(0.02)
    status, timed_out = await relay.result(
      timeout_submit["request_id"],
      gateway_session_id="session-a",
      user_id="alice",
    )
    assert status == "ok"
    assert timed_out["state"] == "timeout"
    assert timed_out["error"]["message"] == "Request timed out"

  run_async(scenario())


def test_chat_result_poll_uses_submitter_session_when_delivery_session_differs() -> None:
  async def scenario() -> None:
    relay = McpRelay()
    client = await relay.register_client("taskpane-session", "A.xlsx", user_id="alice")

    submitted = await relay.submit(
      "send_chat_message",
      {"text": "split owner"},
      timeout=1,
      gateway_session_id="mcp-session",
      user_id="alice",
      kind="chat",
    )
    event = await _event(client.queue)
    entry = relay._inflight[submitted["request_id"]]

    assert event["request_id"] == submitted["request_id"]
    assert entry["session_id"] == "taskpane-session"
    assert entry["user_id"] == "alice"
    assert entry["owner_session_id"] == "mcp-session"
    assert entry["owner_user_id"] == "alice"
    assert await relay.ack(
      event["request_id"],
      event["nonce"],
      event["delivery_id"],
      gateway_session_id="mcp-session",
      user_id="alice",
    ) == "session_mismatch"
    assert await relay.ack(
      event["request_id"],
      event["nonce"],
      event["delivery_id"],
      gateway_session_id="taskpane-session",
      user_id="alice",
    ) == "ok"
    assert await relay.complete(
      event["request_id"],
      event["nonce"],
      event["delivery_id"],
      {"assistant_text": "completed by taskpane"},
      None,
      gateway_session_id="taskpane-session",
      user_id="alice",
    ) == "ok"

    status, done = await relay.result(
      submitted["request_id"],
      gateway_session_id="mcp-session",
      user_id="alice",
    )

    assert status == "ok"
    assert done == {"state": "done", "result": {"assistant_text": "completed by taskpane"}}

  run_async(scenario())


def test_chat_result_poll_rejects_non_submitter_identity() -> None:
  async def scenario() -> None:
    relay = McpRelay()
    client = await relay.register_client("taskpane-session", "A.xlsx", user_id="alice")

    submitted = await relay.submit(
      "send_chat_message",
      {"text": "private result"},
      timeout=1,
      gateway_session_id="mcp-session",
      user_id="alice",
      kind="chat",
    )
    event = await _event(client.queue)
    assert await relay.ack(
      event["request_id"],
      event["nonce"],
      event["delivery_id"],
      gateway_session_id="taskpane-session",
      user_id="alice",
    ) == "ok"
    assert await relay.complete(
      event["request_id"],
      event["nonce"],
      event["delivery_id"],
      {"assistant_text": "secret"},
      None,
      gateway_session_id="taskpane-session",
      user_id="alice",
    ) == "ok"

    status, payload = await relay.result(
      submitted["request_id"],
      gateway_session_id="other-mcp-session",
      user_id="alice",
    )
    assert status == "session_mismatch"
    assert payload == {}

    status, payload = await relay.result(
      submitted["request_id"],
      gateway_session_id="mcp-session",
      user_id="mallory",
    )
    assert status == "user_mismatch"
    assert payload == {}

  run_async(scenario())
