from __future__ import annotations

import asyncio
from typing import Any, Dict, Optional

import pytest

from excel_mcp.relay import McpRelay
from excel_mcp.tool_registry import get_tool_specs


def run_async(coro):  # type: ignore[no-untyped-def]
  return asyncio.run(coro)


def make_relay() -> McpRelay:
  relay = McpRelay()
  relay.ACK_REPLAY_GRACE_SECONDS = 0.05
  return relay


async def register_client(
  relay: McpRelay,
  session: str,
  workbook_name: str,
  *,
  user_id: str = "alice",
) -> Any:
  return await relay.register_client(session, workbook_name, user_id=user_id)


async def get_event(queue: "asyncio.Queue[Dict[str, Any]]", timeout: float = 0.25) -> Dict[str, Any]:
  return await asyncio.wait_for(queue.get(), timeout=timeout)


async def assert_queue_empty(queue: "asyncio.Queue[Dict[str, Any]]", timeout: float = 0.05) -> None:
  try:
    event = await asyncio.wait_for(queue.get(), timeout=timeout)
  except asyncio.TimeoutError:
    return
  raise AssertionError(f"Expected queue to stay empty, got {event}")


async def finish_request(relay: McpRelay, event: Dict[str, Any]) -> Dict[str, Any]:
  assert await relay.ack(event["request_id"], event["nonce"], event["delivery_id"]) == "ok"
  assert await relay.complete(
    event["request_id"],
    event["nonce"],
    event["delivery_id"],
    {"ok": True},
    None,
  ) == "ok"
  return {"ok": True}


def workbook_by_session(workbooks: list[Dict[str, Any]], session: str) -> Dict[str, Any]:
  matches = [workbook for workbook in workbooks if workbook["session"] == session]
  assert len(matches) == 1
  return matches[0]


def assert_workbook(
  workbook: Dict[str, Any],
  *,
  name: str,
  session: str,
  gateway_session_id: str,
  active: bool,
  connected_at: float,
  detached: bool = False,
  detach_grace_deadline: Optional[float] = None,
) -> None:
  assert workbook["name"] == name
  assert workbook["session"] == session
  assert workbook["gateway_session_id"] == gateway_session_id
  assert workbook["active"] is active
  assert workbook["connected_at"] == connected_at
  assert workbook["detached"] is detached
  assert workbook["detach_grace_deadline"] == detach_grace_deadline


def workbook_activity(workbooks: list[Dict[str, Any]]) -> Dict[str, bool]:
  return {workbook["session"]: workbook["active"] for workbook in workbooks}


def test_workbook_tool_schemas_only_inject_workbook_override_for_workbook_bound_tools() -> None:
  specs = {spec["name"]: spec for spec in get_tool_specs()}

  assert "_workbook" in specs["read_cells"]["input_schema"]["properties"]
  assert "_workbook" not in specs["list_workbooks"]["input_schema"]["properties"]
  assert "_workbook" not in specs["switch_active_workbook"]["input_schema"]["properties"]


def test_list_workbooks_reports_zero_one_and_many_connections() -> None:
  async def scenario() -> None:
    relay = make_relay()
    assert await relay.list_workbooks() == []

    first = await register_client(relay, "session-a", "A.xlsx")
    listed = await relay.list_workbooks()
    assert len(listed) == 1
    assert_workbook(
      listed[0],
      name="A.xlsx",
      session=first.workbook_session,
      gateway_session_id=first.gateway_session_id,
      active=True,
      connected_at=first.connected_at,
    )

    second = await register_client(relay, "session-b", "B.xlsx")
    listed = await relay.list_workbooks()
    assert len(listed) == 2
    assert_workbook(
      workbook_by_session(listed, first.workbook_session),
      name="A.xlsx",
      session=first.workbook_session,
      gateway_session_id=first.gateway_session_id,
      active=True,
      connected_at=first.connected_at,
    )
    assert_workbook(
      workbook_by_session(listed, second.workbook_session),
      name="B.xlsx",
      session=second.workbook_session,
      gateway_session_id=second.gateway_session_id,
      active=True,
      connected_at=second.connected_at,
    )

  run_async(scenario())


def test_switch_active_workbook_flips_active_session_and_emits_event() -> None:
  async def scenario() -> None:
    relay = make_relay()
    await register_client(relay, "session-a", "A.xlsx")
    second = await register_client(relay, "session-b", "B.xlsx")

    response = await relay.switch_active_workbook("session-b")
    assert response["status"] == "ok"
    assert response["active"] == "session-b"
    listed = await relay.list_workbooks()
    assert workbook_by_session(listed, second.workbook_session)["active"] is True
    event = await get_event(second.queue)
    assert event["type"] == "active_changed"
    assert event["new_active"] == "session-b"

  run_async(scenario())


def test_switch_active_workbook_rejects_unknown_session() -> None:
  async def scenario() -> None:
    relay = make_relay()
    await register_client(relay, "session-a", "A.xlsx")

    with pytest.raises(RuntimeError, match="unknown_session"):
      await relay.switch_active_workbook("missing")

  run_async(scenario())


def test_workbook_override_routes_to_target_without_changing_active_session() -> None:
  async def scenario() -> None:
    relay = make_relay()
    first = await register_client(relay, "session-a", "A.xlsx")
    second = await register_client(relay, "session-b", "B.xlsx")
    active_before = workbook_activity(await relay.list_workbooks())

    execute_task = asyncio.create_task(
      relay.execute("read_cells", {"range": "A1", "_workbook": "session-b"}, timeout=1)
    )
    event = await get_event(second.queue)
    await assert_queue_empty(first.queue)
    assert workbook_activity(await relay.list_workbooks()) == active_before

    await finish_request(relay, event)
    result = await execute_task
    assert result["error"] is None
    assert workbook_activity(await relay.list_workbooks()) == active_before

  run_async(scenario())


def test_workbook_override_rejects_unknown_session_without_fallback() -> None:
  async def scenario() -> None:
    relay = make_relay()
    first = await register_client(relay, "session-a", "A.xlsx")
    active_before = workbook_activity(await relay.list_workbooks())

    with pytest.raises(RuntimeError, match="unknown_session"):
      await relay.execute("read_cells", {"range": "A1", "_workbook": "missing"}, timeout=1)

    assert workbook_activity(await relay.list_workbooks()) == active_before
    listed = await relay.list_workbooks()
    assert workbook_by_session(listed, first.workbook_session)["active"] is True
    await assert_queue_empty(first.queue)

  run_async(scenario())


def test_workbook_override_is_stripped_before_forwarding_to_sse() -> None:
  async def scenario() -> None:
    relay = make_relay()
    await register_client(relay, "session-a", "A.xlsx")
    second = await register_client(relay, "session-b", "B.xlsx")

    execute_task = asyncio.create_task(
      relay.execute("read_cells", {"range": "A1", "_workbook": "session-b"}, timeout=1)
    )
    event = await get_event(second.queue)
    assert event["tool_input"] == {"range": "A1"}

    await finish_request(relay, event)
    result = await execute_task
    assert result["error"] is None

  run_async(scenario())
