"""Sending must always be confirmed by the user."""

from types import SimpleNamespace

import pytest

from m365_mcp import server

DRAFT = {"id": "d1", "subject": "Hello", "isDraft": True, "hasAttachments": False,
         "toRecipients": [{"emailAddress": {"address": "a@b.com"}}], "ccRecipients": [], "bccRecipients": []}


class FakeCtx:
    def __init__(self, caps=None, answer=None):
        self.client_capabilities = caps
        self._answer = answer
        self.asked = None

    async def elicit(self, message, schema):
        self.asked = message
        return self._answer


@pytest.fixture(autouse=True)
def use_test_graph(graph, monkeypatch):
    monkeypatch.setattr(server, "_graph", graph)
    monkeypatch.setattr(server.cfg, "send_mode", "confirm")


async def test_without_prompt_support_it_asks_in_conversation(mock):
    mock.get("/me/messages/d1").respond(json=DRAFT)
    send = mock.post("/me/messages/d1/send").respond(status_code=202)
    out = await server.send_drafts(["d1"], FakeCtx())
    assert out["needs_confirmation"] and "a@b.com" in out["summary"]
    assert not send.called

    out = await server.send_drafts(["d1"], FakeCtx(), user_confirmed=True)
    assert out["results"] == [{"draft_id": "d1", "sent": True}]


async def test_prompt_declined_sends_nothing(mock):
    mock.get("/me/messages/d1").respond(json=DRAFT)
    send = mock.post("/me/messages/d1/send").respond(status_code=202)
    caps = SimpleNamespace(elicitation=SimpleNamespace(form=SimpleNamespace(), url=None))
    ctx = FakeCtx(caps, SimpleNamespace(action="decline", data=None))
    # user_confirmed from the model can't override a prompt the user declined
    out = await server.send_drafts(["d1"], ctx, user_confirmed=True)
    assert out["sent"] is False and not send.called
    assert "Hello" in ctx.asked


async def test_prompt_accepted_sends_all(mock):
    for d in ("d1", "d2"):
        mock.get(f"/me/messages/{d}").respond(json={**DRAFT, "id": d})
    s1 = mock.post("/me/messages/d1/send").respond(status_code=202)
    s2 = mock.post("/me/messages/d2/send").respond(status_code=202)
    caps = SimpleNamespace(elicitation=SimpleNamespace(form=SimpleNamespace(), url=None))
    ctx = FakeCtx(caps, SimpleNamespace(action="accept", data=SimpleNamespace(send=True)))
    out = await server.send_drafts(["d1", "d2"], ctx)
    assert s1.called and s2.called and len(out["results"]) == 2
    assert ctx.asked.startswith("Send 2 emails?")


async def test_sending_can_be_disabled(mock, monkeypatch):
    monkeypatch.setattr(server.cfg, "send_mode", "disabled")
    with pytest.raises(Exception, match="turned off"):
        await server.send_drafts(["d1"], FakeCtx(), user_confirmed=True)


async def test_refuses_to_send_non_drafts(mock):
    mock.get("/me/messages/d1").respond(json={**DRAFT, "isDraft": False})
    with pytest.raises(Exception, match="not a draft"):
        await server.send_drafts(["d1"], FakeCtx(), user_confirmed=True)
