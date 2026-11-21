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

    out = await server.send_drafts(["d1"], FakeCtx(), confirmation_code=out["confirmation_code"])
    assert out["results"] == [{"draft_id": "d1", "sent": True}]


async def test_prompt_declined_sends_nothing(mock):
    mock.get("/me/messages/d1").respond(json=DRAFT)
    send = mock.post("/me/messages/d1/send").respond(status_code=202)
    caps = SimpleNamespace(elicitation=SimpleNamespace(form=SimpleNamespace(), url=None))
    ctx = FakeCtx(caps, SimpleNamespace(action="decline", data=None))
    # a code from the model can't override a prompt the user declined
    out = await server.send_drafts(["d1"], ctx, confirmation_code="ABCDEF")
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
        await server.send_drafts(["d1"], FakeCtx(), confirmation_code="ABCDEF")


async def test_refuses_to_send_non_drafts(mock):
    mock.get("/me/messages/d1").respond(json={**DRAFT, "isDraft": False})
    with pytest.raises(Exception, match="not a draft"):
        await server.send_drafts(["d1"], FakeCtx(), confirmation_code="ABCDEF")


async def test_sending_needs_a_code_from_a_summary(mock):
    mock.get("/me/messages/d1").respond(json=DRAFT)
    send = mock.post("/me/messages/d1/send").respond(status_code=202)
    with pytest.raises(Exception, match="isn't valid"):
        await server.send_drafts(["d1"], FakeCtx(), confirmation_code="MADEUP")
    assert not send.called


async def test_a_code_works_once(mock):
    mock.get("/me/messages/d1").respond(json=DRAFT)
    send = mock.post("/me/messages/d1/send").respond(status_code=202)
    code = (await server.send_drafts(["d1"], FakeCtx()))["confirmation_code"]
    await server.send_drafts(["d1"], FakeCtx(), confirmation_code=code)
    with pytest.raises(Exception, match="isn't valid"):
        await server.send_drafts(["d1"], FakeCtx(), confirmation_code=code)
    assert send.call_count == 1


async def test_a_code_doesnt_cover_other_drafts(mock):
    mock.get("/me/messages/d1").respond(json=DRAFT)
    mock.get("/me/messages/d2").respond(json={**DRAFT, "id": "d2"})
    send = mock.post(url__regex=r".*/send$").respond(status_code=202)
    code = (await server.send_drafts(["d1"], FakeCtx()))["confirmation_code"]
    with pytest.raises(Exception, match="aren't the same items"):
        await server.send_drafts(["d2"], FakeCtx(), confirmation_code=code)
    assert not send.called


async def test_a_code_fails_if_recipients_changed_after_the_summary(mock):
    route = mock.get("/me/messages/d1").respond(json=DRAFT)
    send = mock.post("/me/messages/d1/send").respond(status_code=202)
    code = (await server.send_drafts(["d1"], FakeCtx()))["confirmation_code"]
    route.respond(json={**DRAFT, "toRecipients": [{"emailAddress": {"address": "someone.else@x.com"}}]})
    with pytest.raises(Exception, match="changed since the summary"):
        await server.send_drafts(["d1"], FakeCtx(), confirmation_code=code)
    assert not send.called


async def test_a_failed_prompt_falls_back_to_a_code(mock):
    class BrokenCtx(FakeCtx):
        async def elicit(self, message, schema):
            raise RuntimeError("client can't show it")
    mock.get("/me/messages/d1").respond(json=DRAFT)
    send = mock.post("/me/messages/d1/send").respond(status_code=202)
    caps = SimpleNamespace(elicitation=SimpleNamespace(form=SimpleNamespace(), url=None))
    out = await server.send_drafts(["d1"], BrokenCtx(caps))
    assert out["needs_confirmation"] and out["confirmation_code"] and not send.called


def test_expired_codes_are_refused(monkeypatch):
    from m365_mcp import confirm
    c = confirm.Confirmations(ttl=-1)
    code = c.issue("send", "fp")
    with pytest.raises(ValueError, match="expired"):
        c.redeem("send", code, "fp")
