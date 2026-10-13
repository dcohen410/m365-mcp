import json

import pytest

from m365_mcp import events

TZ = "AUS Eastern Standard Time"


@pytest.fixture
def tz(mock):
    return mock.get("/me/mailboxSettings").respond(json={"timeZone": TZ})


def event(**over):
    out = {"id": "e1", "subject": "Leave", "isAllDay": True, "showAs": "free",
           "start": {"dateTime": "2026-11-09T00:00:00.0000000", "timeZone": TZ},
           "end": {"dateTime": "2026-11-27T00:00:00.0000000", "timeZone": TZ}, "attendees": []}
    out.update(over)
    return out


async def test_all_day_event_runs_to_midnight_after_last_day(mock, graph, tz):
    route = mock.post("/me/events").respond(json=event())
    out = await events.create_event(graph, "Leave", "2026-11-09", "2026-11-26", all_day=True, show_as="Free")
    sent = json.loads(route.calls[0].request.content)
    assert sent["isAllDay"] is True and sent["showAs"] == "free" and "attendees" not in sent
    assert sent["start"] == {"dateTime": "2026-11-09T00:00:00", "timeZone": TZ}
    assert sent["end"] == {"dateTime": "2026-11-27T00:00:00", "timeZone": TZ}
    assert out["start"] == "2026-11-09" and out["end"] == "2026-11-26"


async def test_single_all_day_event_without_end(mock, graph, tz):
    route = mock.post("/me/events").respond(json=event())
    await events.create_event(graph, "Leave", "2026-10-16", all_day=True)
    sent = json.loads(route.calls[0].request.content)
    assert sent["end"]["dateTime"] == "2026-10-17T00:00:00"


async def test_timed_event_and_reminder_off(mock, graph, tz):
    route = mock.post("/me/events").respond(json=event(isAllDay=False))
    await events.create_event(graph, "Call", "2026-10-16T09:30", "2026-10-16T10:00", reminder_minutes=-1)
    sent = json.loads(route.calls[0].request.content)
    assert sent["start"]["dateTime"] == "2026-10-16T09:30:00" and sent["isReminderOn"] is False


async def test_bad_input_is_refused(graph):
    with pytest.raises(ValueError, match="show_as"):
        events._check_show_as("away")
    with pytest.raises(ValueError, match="before the first"):
        events._times("2026-10-16", "2026-10-15", True, TZ)
    with pytest.raises(ValueError, match="needs an `end`"):
        events._times("2026-10-16T09:00", None, False, TZ)


async def test_list_uses_mailbox_time_zone_and_filters(mock, graph, tz):
    route = mock.get("/me/calendarView").respond(json={"value": [event(), event(id="e2", subject="Dentist")]})
    out = await events.list_events(graph, "2026-11-01", "2026-11-30", query="dentist")
    req = route.calls[0].request
    assert req.headers["Prefer"] == f'outlook.timezone="{TZ}"'
    assert req.url.params["endDateTime"] == "2026-12-01T00:00:00"
    assert [e["id"] for e in out] == ["e2"]


async def test_update_refuses_events_with_attendees(mock, graph):
    mock.get("/me/events/e1").respond(json=event(attendees=[{"emailAddress": {"address": "a@b.com"}}]))
    with pytest.raises(ValueError, match="attendees"):
        await events.update_event(graph, "e1", show_as="busy")


async def test_update_turns_timed_event_into_all_day(mock, graph, tz):
    mock.get("/me/events/e1").respond(json=event(isAllDay=False, start={"dateTime": "2026-10-16T00:00:00", "timeZone": TZ},
                                                 end={"dateTime": "2026-10-17T00:00:00", "timeZone": TZ}))
    route = mock.patch("/me/events/e1").respond(json=event())
    await events.update_event(graph, "e1", start="2026-10-16", all_day=True)
    sent = json.loads(route.calls[0].request.content)
    assert sent["isAllDay"] is True
    assert sent["end"]["dateTime"] == "2026-10-17T00:00:00"


async def test_update_show_as_only_leaves_times_alone(mock, graph):
    mock.get("/me/events/e1").respond(json=event())
    route = mock.patch("/me/events/e1").respond(json=event(showAs="oof"))
    await events.update_event(graph, "e1", show_as="oof")
    assert json.loads(route.calls[0].request.content) == {"showAs": "oof"}
