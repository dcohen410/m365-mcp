"""Deleting files must always be confirmed, and only ever sends files to the recycle bin."""

from types import SimpleNamespace

import pytest

from m365_mcp import server
from tests.conftest import item
from tests.test_send import FakeCtx

FORM = SimpleNamespace(elicitation=SimpleNamespace(form=SimpleNamespace(), url=None))


@pytest.fixture(autouse=True)
def use_test_graph(graph, monkeypatch):
    monkeypatch.setattr(server, "_graph", graph)
    monkeypatch.setattr(server.cfg, "delete_mode", "confirm")


async def test_without_prompt_support_it_asks_in_conversation(mock):
    mock.get("/me/drive/root:/Docs/old.pdf:").respond(json=item("old.pdf"))
    delete = mock.delete("/drives/DRIVE/items/ITEM").respond(204)
    out = await server.delete_files(["Docs/old.pdf"], FakeCtx())
    assert out["needs_confirmation"] and "old.pdf" in out["summary"] and "recycle bin" in out["summary"]
    assert not delete.called
    out = await server.delete_files(["Docs/old.pdf"], FakeCtx(), confirmation_code=out["confirmation_code"])
    assert out["results"] == [{"name": "old.pdf", "deleted": True}] and delete.called


async def test_declined_prompt_overrides_model_flag(mock):
    mock.get("/me/drive/root:/Docs/old.pdf:").respond(json=item("old.pdf"))
    delete = mock.delete("/drives/DRIVE/items/ITEM").respond(204)
    ctx = FakeCtx(FORM, SimpleNamespace(action="decline", data=None))
    out = await server.delete_files(["Docs/old.pdf"], ctx, confirmation_code="ABCDEF")
    assert out["deleted"] is False and not delete.called


async def test_accepted_prompt_deletes(mock):
    mock.get("/me/drive/root:/Docs/old.pdf:").respond(json=item("old.pdf"))
    delete = mock.delete("/drives/DRIVE/items/ITEM").respond(204)
    ctx = FakeCtx(FORM, SimpleNamespace(action="accept", data=SimpleNamespace(delete=True)))
    out = await server.delete_files(["Docs/old.pdf"], ctx)
    assert delete.called and out["results"][0]["deleted"]
    assert ctx.asked.startswith("Delete 1 file?")


async def test_folders_with_contents_are_refused(mock):
    full = item("Docs", folder=True)
    full["folder"]["childCount"] = 3
    mock.get("/me/drive/root:/Docs:").respond(json=full)
    delete = mock.delete("/drives/DRIVE/items/ITEM").respond(204)
    with pytest.raises(Exception, match="is a folder with 3 item"):
        await server.delete_files(["Docs"], FakeCtx(), confirmation_code="ABCDEF")
    assert not delete.called


async def test_empty_folders_need_confirmation_then_delete(mock):
    mock.get("/me/drive/root:/Docs/Old:").respond(json=item("Old", folder=True))
    delete = mock.delete("/drives/DRIVE/items/ITEM").respond(204)
    out = await server.delete_files(["Docs/Old"], FakeCtx())
    assert out["needs_confirmation"] and "1 empty folder" in out["summary"] and not delete.called
    out = await server.delete_files(["Docs/Old"], FakeCtx(), confirmation_code=out["confirmation_code"])
    assert out["results"] == [{"name": "Old", "deleted": True}] and delete.called


async def test_deleting_can_be_disabled(monkeypatch):
    monkeypatch.setattr(server.cfg, "delete_mode", "disabled")
    with pytest.raises(Exception, match="turned off"):
        await server.delete_files(["Docs/old.pdf"], FakeCtx(), confirmation_code="ABCDEF")


def test_disabled_deleting_hides_the_tool():
    from m365_mcp.config import Config
    tools = server.registered_tools()
    assert "delete_files" in server.hidden_tools(Config(delete_mode="disabled"), tools)
    assert "delete_files" not in server.hidden_tools(Config(), tools)
    assert "delete_files" in server.hidden_tools(Config(read_only=True), tools)


async def test_delete_code_survives_a_changing_download_link(mock):
    first, second = item("old.pdf"), item("old.pdf")
    first["@microsoft.graph.downloadUrl"], second["@microsoft.graph.downloadUrl"] = "https://x/1", "https://x/2"
    route = mock.get("/me/drive/root:/Docs/old.pdf:").respond(json=first)
    delete = mock.delete("/drives/DRIVE/items/ITEM").respond(204)
    code = (await server.delete_files(["Docs/old.pdf"], FakeCtx()))["confirmation_code"]
    route.respond(json=second)
    await server.delete_files(["Docs/old.pdf"], FakeCtx(), confirmation_code=code)
    assert delete.called


async def test_delete_code_doesnt_cover_other_files(mock):
    mock.get("/me/drive/root:/Docs/old.pdf:").respond(json=item("old.pdf"))
    mock.get("/me/drive/root:/Docs/keep.pdf:").respond(json=item("keep.pdf", item_id="KEEP"))
    delete = mock.delete(url__regex=r".*/items/.*").respond(204)
    code = (await server.delete_files(["Docs/old.pdf"], FakeCtx()))["confirmation_code"]
    with pytest.raises(Exception, match="aren't the same items"):
        await server.delete_files(["Docs/keep.pdf"], FakeCtx(), confirmation_code=code)
    assert not delete.called
