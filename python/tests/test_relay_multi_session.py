from __future__ import annotations

import asyncio
from typing import Any, Dict

import pytest

from excel_mcp.relay import McpRelay


def run_async(coro):  # type: ignore[no-untyped-def]
  return asyncio.run(coro)


def make_relay(grace: float = 0.05) -> McpRelay:
  relay = McpRelay()
  relay.ACK_REPLAY_GRACE_SECONDS = grace
  relay.reconnect_grace_seconds = grace
  return relay


async def register_client(
  relay: McpRelay,
  session: str | None,
  workbook_name: str | None,
  *,
  user_id: str = "alice",
  **kwargs: Any,
):
  return await relay.register_client(session, workbook_name, user_id=user_id, **kwargs)


async def get_event(queue: "asyncio.Queue[Dict[str, Any]]", timeout: float = 0.25) -> Dict[str, Any]:
  return await asyncio.wait_for(queue.get(), timeout=timeout)


async def assert_queue_empty(queue: "asyncio.Queue[Dict[str, Any]]", timeout: float = 0.05) -> None:
  try:
    event = await asyncio.wait_for(queue.get(), timeout=timeout)
  except asyncio.TimeoutError:
    return
  raise AssertionError(f"Expected queue to stay empty, got {event}")


async def finish_request(
  relay: McpRelay,
  event: Dict[str, Any],
  *,
  result: Dict[str, Any] | None = None,
  delivery_id: str | None = None,
) -> str:
  ack_status = await relay.ack(
    event["request_id"],
    event["nonce"],
    delivery_id if delivery_id is not None else event.get("delivery_id"),
  )
  assert ack_status == "ok"
  complete_status = await relay.complete(
    event["request_id"],
    event["nonce"],
    delivery_id if delivery_id is not None else event.get("delivery_id"),
    result or {"ok": True},
    None,
  )
  assert complete_status == "ok"
  return event["request_id"]


def test_1_registers_two_sessions() -> None:
  async def scenario() -> None:
    relay = make_relay()
    a = await register_client(relay, "session-a", "A.xlsx")
    b = await register_client(relay, "session-b", "B.xlsx")

    assert set(relay._clients) == {"session-a", "session-b"}
    assert relay._clients["session-a"].client_id == a.client_id
    assert relay._clients["session-b"].client_id == b.client_id

  run_async(scenario())


def test_2_explicit_target_routes_only_to_named_session() -> None:
  async def scenario() -> None:
    relay = make_relay()
    a = await register_client(relay, "session-a", "A.xlsx")
    b = await register_client(relay, "session-b", "B.xlsx")

    execute_task = asyncio.create_task(
      relay.execute("read_cells", {"range": "A1"}, timeout=1, target_session="session-a")
    )
    event = await get_event(a.queue)
    assert event["tool_name"] == "read_cells"
    assert event["delivery_id"] == a.client_id
    await assert_queue_empty(b.queue)

    await finish_request(relay, event)
    result = await execute_task
    assert result["error"] is None

  run_async(scenario())


def test_3_active_workbook_is_scoped_to_gateway_session() -> None:
  async def scenario() -> None:
    relay = make_relay()
    a = await register_client(relay, "session-a", "A.xlsx")
    b = await register_client(relay, "session-b", "B.xlsx")

    assert relay._active_workbook == {"session-a": "session-a", "session-b": "session-b"}

    execute_task = asyncio.create_task(relay.execute("read_cells", {"range": "A1"}, timeout=1))
    event = await get_event(a.queue)
    assert event["delivery_id"] == a.client_id
    await assert_queue_empty(b.queue)

    await finish_request(relay, event)
    result = await execute_task
    assert result["error"] is None

  run_async(scenario())


def test_4_same_session_reconnect_replaces_old_slot() -> None:
  async def scenario() -> None:
    relay = make_relay()
    first = await register_client(relay, "shared", "A.xlsx")
    second = await register_client(relay, "shared", "A.xlsx")

    replaced = await get_event(first.queue)
    assert replaced == {"type": "replaced", "reason": "Same-session reconnect"}

    await relay.unregister_client("shared", first.client_id)
    assert relay._clients["shared"].client_id == second.client_id

  run_async(scenario())


def test_5_stale_ack_is_rejected_after_replay() -> None:
  async def scenario() -> None:
    relay = make_relay()
    first = await register_client(relay, "session-a", "A.xlsx")
    execute_task = asyncio.create_task(relay.execute("read_cells", {"range": "A1"}, timeout=1))
    original = await get_event(first.queue)

    second = await register_client(relay, "session-a", "A.xlsx")
    replay = await get_event(second.queue)
    assert replay["replay"] is True

    status = await relay.ack(original["request_id"], original["nonce"], original["delivery_id"])
    assert status == "stale_delivery"
    assert relay._inflight[original["request_id"]]["state"] == "delivered"

    await finish_request(relay, replay)
    result = await execute_task
    assert result["error"] is None

  run_async(scenario())


def test_6_stale_complete_is_rejected_after_replay() -> None:
  async def scenario() -> None:
    relay = make_relay()
    first = await register_client(relay, "session-a", "A.xlsx")
    execute_task = asyncio.create_task(relay.execute("read_cells", {"range": "A1"}, timeout=1))
    original = await get_event(first.queue)

    second = await register_client(relay, "session-a", "A.xlsx")
    replay = await get_event(second.queue)
    assert replay["replay"] is True

    status = await relay.complete(
      original["request_id"],
      original["nonce"],
      original["delivery_id"],
      {"bad": True},
      None,
    )
    assert status == "stale_delivery"
    assert relay._inflight[original["request_id"]]["state"] == "delivered"

    await finish_request(relay, replay)
    result = await execute_task
    assert result["error"] is None

  run_async(scenario())


def test_6a_missing_delivery_id_is_rejected_for_modern_ack() -> None:
  async def scenario() -> None:
    relay = make_relay()
    client = await register_client(relay, "session-a", "A.xlsx")
    execute_task = asyncio.create_task(relay.execute("read_cells", {"range": "A1"}, timeout=1))
    event = await get_event(client.queue)

    status = await relay.ack(event["request_id"], event["nonce"], None)
    assert status == "delivery_id_required"
    assert relay._inflight[event["request_id"]]["state"] == "delivered"

    await finish_request(relay, event)
    result = await execute_task
    assert result["error"] is None

  run_async(scenario())


def test_6b_missing_delivery_id_is_rejected_for_all_sessions() -> None:
  async def scenario() -> None:
    relay = make_relay()
    client = await register_client(relay, "session-a", "A.xlsx")
    execute_task = asyncio.create_task(
      relay.execute("read_cells", {"range": "A1"}, timeout=1, target_session=client.gateway_session_id)
    )
    event = await get_event(client.queue)

    ack_status = await relay.ack(event["request_id"], event["nonce"], None)
    assert ack_status == "delivery_id_required"

    await finish_request(relay, event)
    result = await execute_task
    assert result["error"] is None

  run_async(scenario())


def test_6c_request_scoped_delivery_id_blocks_old_handler_after_reconnect() -> None:
  async def scenario() -> None:
    relay = make_relay()
    first = await register_client(relay, "session-a", "A.xlsx")
    execute_task = asyncio.create_task(relay.execute("write_cells", {"range": "A1", "values": 1}, timeout=1))
    original = await get_event(first.queue)

    second = await register_client(relay, "session-a", "A.xlsx")
    replay = await get_event(second.queue)
    executed: list[str] = []

    async def run_handler(event: Dict[str, Any]) -> str:
      ack_status = await relay.ack(event["request_id"], event["nonce"], event["delivery_id"])
      if ack_status != "ok":
        return ack_status
      executed.append(event["delivery_id"])
      complete_status = await relay.complete(
        event["request_id"],
        event["nonce"],
        event["delivery_id"],
        {"ok": True},
        None,
      )
      assert complete_status == "ok"
      return ack_status

    stale_status = await run_handler(original)
    assert stale_status == "stale_delivery"
    assert executed == []

    fresh_status = await run_handler(replay)
    assert fresh_status == "ok"
    assert executed == [second.client_id]

    result = await execute_task
    assert result["result"] == {"ok": True}

  run_async(scenario())


def test_6d_rapid_reconnects_only_replay_once() -> None:
  async def scenario() -> None:
    relay = make_relay()
    first = await register_client(relay, "session-a", "A.xlsx")
    execute_task = asyncio.create_task(relay.execute("read_cells", {"range": "A1"}, timeout=1))
    initial = await get_event(first.queue)

    await relay.unregister_client("session-a", first.client_id)
    await register_client(relay, "session-a", "A.xlsx")
    await asyncio.sleep(relay.ACK_REPLAY_GRACE_SECONDS / 2)
    second = relay._clients["session-a"]
    await relay.unregister_client("session-a", second.client_id)
    third = await register_client(relay, "session-a", "A.xlsx")

    await asyncio.sleep(relay.ACK_REPLAY_GRACE_SECONDS + 0.03)
    replay = await get_event(third.queue)
    assert replay["type"] == "mcp_tool_request"
    assert replay["delivery_id"] == third.client_id
    await assert_queue_empty(third.queue)
    assert relay._inflight[initial["request_id"]]["delivery_client_id"] == third.client_id

    await finish_request(relay, replay)
    result = await execute_task
    assert result["error"] is None

  run_async(scenario())


def test_6e_legacy_namespace_and_missing_session_are_rejected() -> None:
  async def scenario() -> None:
    relay = make_relay()

    with pytest.raises(ValueError, match="legacy session tokens"):
      await register_client(relay, "legacy:fake", "X.xlsx")
    with pytest.raises(ValueError, match="gateway_session_id or session is required"):
      await register_client(relay, None, "Y.xlsx")

    assert relay._clients == {}

  run_async(scenario())


def test_6ea_registration_requires_user_id() -> None:
  async def scenario() -> None:
    relay = make_relay()

    with pytest.raises(ValueError, match="user_id is required"):
      await relay.register_client("session-a", "A.xlsx")

  run_async(scenario())


def test_6f_modern_sessions_cannot_bypass_delivery_id_requirement() -> None:
  async def scenario() -> None:
    relay = make_relay()
    client = await register_client(relay, "session-a", "A.xlsx")
    execute_task = asyncio.create_task(relay.execute("read_cells", {"range": "A1"}, timeout=1))
    event = await get_event(client.queue)

    status = await relay.ack(event["request_id"], event["nonce"], None)
    assert status == "delivery_id_required"
    assert "is_legacy" not in relay._inflight[event["request_id"]]
    assert relay._inflight[event["request_id"]]["state"] == "delivered"

    await finish_request(relay, event)
    result = await execute_task
    assert result["error"] is None

  run_async(scenario())


def test_7_same_session_reconnect_within_grace_replays_after_delay() -> None:
  async def scenario() -> None:
    relay = make_relay()
    first = await register_client(relay, "session-a", "A.xlsx")
    execute_task = asyncio.create_task(relay.execute("read_cells", {"range": "A1"}, timeout=1))
    initial = await get_event(first.queue)

    await relay.unregister_client("session-a", first.client_id)
    second = await register_client(relay, "session-a", "A.xlsx")

    await assert_queue_empty(second.queue, timeout=0.02)
    assert relay._inflight[initial["request_id"]]["delivery_client_id"] == first.client_id

    replay = await get_event(second.queue)
    assert replay["replay"] is True
    assert replay["delivery_id"] == second.client_id

    await finish_request(relay, replay)
    result = await execute_task
    assert result["error"] is None

  run_async(scenario())


def test_8_same_session_reconnect_after_grace_does_not_replay_inflight() -> None:
  async def scenario() -> None:
    relay = make_relay()
    first = await register_client(relay, "session-a", "A.xlsx")
    execute_task = asyncio.create_task(relay.execute("read_cells", {"range": "A1"}, timeout=0.2))
    await get_event(first.queue)

    await relay.unregister_client("session-a", first.client_id)
    await asyncio.sleep(relay.reconnect_grace_seconds + 0.02)
    second = await register_client(relay, "session-a", "A.xlsx")

    await assert_queue_empty(second.queue, timeout=0.1)

    with pytest.raises(asyncio.TimeoutError):
      await execute_task

  run_async(scenario())


def test_9_cross_session_replay_never_falls_over_to_other_session() -> None:
  async def scenario() -> None:
    relay = make_relay()
    a = await register_client(relay, "session-a", "A.xlsx")
    b = await register_client(relay, "session-b", "B.xlsx")

    execute_task = asyncio.create_task(relay.execute("read_cells", {"range": "A1"}, timeout=0.2))
    await get_event(a.queue)
    await relay.unregister_client("session-a", a.client_id)

    with pytest.raises(asyncio.TimeoutError):
      await execute_task
    await assert_queue_empty(b.queue)

  run_async(scenario())


def test_10_unregister_marks_client_detached_without_promoting_other_sessions() -> None:
  async def scenario() -> None:
    relay = make_relay()
    a = await register_client(relay, "session-a", "A.xlsx")
    b = await register_client(relay, "session-b", "B.xlsx")

    await relay.unregister_client("session-a", a.client_id)

    assert relay._clients["session-a"].detached is True
    assert relay._inflight == {}
    await assert_queue_empty(b.queue)

  run_async(scenario())


def test_11_stale_unregister_is_a_noop() -> None:
  async def scenario() -> None:
    relay = make_relay()
    first = await register_client(relay, "session-a", "A.xlsx")
    second = await register_client(relay, "session-a", "A.xlsx")

    await relay.unregister_client("session-a", first.client_id)
    assert relay._clients["session-a"].client_id == second.client_id

  run_async(scenario())


def test_12_missing_session_is_rejected() -> None:
  async def scenario() -> None:
    relay = make_relay()

    with pytest.raises(ValueError, match="gateway_session_id or session is required"):
      await register_client(relay, None, "No Session.xlsx")

  run_async(scenario())


def test_13_multiple_explicit_clients_can_be_targeted_independently() -> None:
  async def scenario() -> None:
    relay = make_relay()
    first = await register_client(relay, "session-a", "A.xlsx")
    second = await register_client(relay, "session-b", "B.xlsx")

    assert first.gateway_session_id != second.gateway_session_id
    assert set(relay._clients) == {first.gateway_session_id, second.gateway_session_id}

    first_task = asyncio.create_task(
      relay.execute("read_cells", {"range": "A1"}, timeout=1, target_session=first.gateway_session_id)
    )
    first_event = await get_event(first.queue)
    await assert_queue_empty(second.queue)
    await finish_request(relay, first_event)
    assert (await first_task)["error"] is None

    second_task = asyncio.create_task(
      relay.execute("read_cells", {"range": "B2"}, timeout=1, target_session=second.gateway_session_id)
    )
    second_event = await get_event(second.queue)
    await finish_request(relay, second_event)
    assert (await second_task)["error"] is None

  run_async(scenario())


def test_14_stale_explicit_target_fails_closed_without_fallback() -> None:
  async def scenario() -> None:
    relay = make_relay()
    stale = await register_client(relay, "session-a", "A.xlsx")
    active = await register_client(relay, "session-b", "B.xlsx")

    await relay.unregister_client(stale.gateway_session_id, stale.client_id)
    await asyncio.sleep(relay.reconnect_grace_seconds + 0.02)

    with pytest.raises(RuntimeError, match="unknown_session"):
      await relay.execute("read_cells", {"range": "A1"}, timeout=1, target_session=stale.gateway_session_id)

    assert relay._active_workbook[active.gateway_session_id] == active.workbook_session
    await assert_queue_empty(active.queue)

  run_async(scenario())


def test_15_delivery_id_cannot_be_omitted() -> None:
  async def scenario() -> None:
    relay = make_relay()

    client = await register_client(relay, "session-a", "A.xlsx")
    task = asyncio.create_task(
      relay.execute("read_cells", {"range": "A1"}, timeout=1, target_session=client.gateway_session_id)
    )
    event = await get_event(client.queue)
    assert await relay.ack(event["request_id"], event["nonce"], None) == "delivery_id_required"

    await finish_request(relay, event)
    assert (await task)["error"] is None

  run_async(scenario())
