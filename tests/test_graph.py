"""Retries must never repeat a request that may already have worked, like sending an email."""

import pytest

from m365_mcp import drives, graph as graph_mod
from m365_mcp.graph import GraphError


@pytest.fixture(autouse=True)
def no_sleep(monkeypatch):
    async def instant(_):
        return None
    monkeypatch.setattr(graph_mod.asyncio, "sleep", instant)


async def test_post_is_not_retried_after_a_gateway_timeout(mock, graph):
    send = mock.post("/me/messages/d1/send").respond(status_code=504, json={"error": {"message": "timeout"}})
    with pytest.raises(GraphError):
        await graph.post("/me/messages/d1/send")
    assert send.call_count == 1


async def test_post_is_retried_when_throttled(mock, graph):
    send = mock.post("/me/messages/d1/send").mock(side_effect=[
        graph_mod.httpx.Response(429, headers={"Retry-After": "1"}), graph_mod.httpx.Response(202)])
    await graph.post("/me/messages/d1/send")
    assert send.call_count == 2


async def test_get_is_retried_after_a_gateway_error(mock, graph):
    route = mock.get("/me").mock(side_effect=[graph_mod.httpx.Response(502), graph_mod.httpx.Response(200, json={})])
    await graph.get("/me")
    assert route.call_count == 2


async def test_retry_after_as_a_date_is_understood(mock, graph):
    route = mock.get("/me").mock(side_effect=[
        graph_mod.httpx.Response(503, headers={"Retry-After": "Wed, 21 Oct 2015 07:28:00 GMT"}),
        graph_mod.httpx.Response(200, json={})])
    await graph.get("/me")
    assert route.call_count == 2


def test_dot_dot_in_drive_paths_is_refused():
    with pytest.raises(ValueError, match=r"\.\."):
        drives.item_ref("/me/drive", "Docs/../../items/ABC")
    assert drives.clean_path("./Docs/./a.pdf") == "Docs/a.pdf"
