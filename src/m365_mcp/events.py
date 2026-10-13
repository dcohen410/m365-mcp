"""Outlook calendar operations on the signed-in person's own calendar.

Events never get attendees here, so nothing creates or sends an invitation.
Times are wall-clock times in the mailbox's own time zone (from Outlook's
mailbox settings), which is how people read their calendar.
"""

from __future__ import annotations

import datetime as dt
import html

from .graph import Graph

SHOW_AS = ("free", "tentative", "busy", "oof", "workingElsewhere")
_EVENT_FIELDS = "id,subject,start,end,isAllDay,showAs,location,bodyPreview,organizer,attendees,isCancelled,webLink"


async def mailbox_time_zone(graph: Graph) -> str:
    """The time zone Outlook shows this person's calendar in, e.g. 'AUS Eastern Standard Time'."""
    settings = await graph.get("/me/mailboxSettings", params={"$select": "timeZone"})
    return settings.get("timeZone") or "UTC"


def _parse_date(value: str, name: str) -> dt.date:
    try:
        return dt.date.fromisoformat(value[:10])
    except ValueError:
        raise ValueError(f"`{name}` must be a date like 2026-10-16, not '{value}'.") from None


def _parse_datetime(value: str, name: str) -> dt.datetime:
    try:
        return dt.datetime.fromisoformat(value)
    except ValueError:
        raise ValueError(f"`{name}` must be a date and time like 2026-10-16T09:30, not '{value}'.") from None


def summarise(e: dict) -> dict:
    start, end = e.get("start") or {}, e.get("end") or {}
    out = {
        "id": e.get("id"),
        "subject": e.get("subject"),
        "all_day": e.get("isAllDay", False),
        "start": start.get("dateTime", "")[:19],
        "end": end.get("dateTime", "")[:19],
        "time_zone": start.get("timeZone"),
        "show_as": e.get("showAs"),
        "location": (e.get("location") or {}).get("displayName") or None,
        "attendees": len(e.get("attendees") or []),
        "cancelled": e.get("isCancelled", False),
        "preview": e.get("bodyPreview") or None,
        "web_link": e.get("webLink"),
    }
    if out["all_day"]:
        # Graph's all-day end is the midnight after the last day; people mean the last day itself.
        out["start"] = out["start"][:10]
        last = dt.date.fromisoformat(out["end"][:10]) - dt.timedelta(days=1) if out["end"] else None
        out["end"] = last.isoformat() if last else ""
    return out


def _times(start: str, end: str | None, all_day: bool, tz: str) -> dict:
    """Graph start/end for an event. For all-day events `end` is the last day (inclusive)."""
    if all_day:
        first = _parse_date(start, "start")
        last = _parse_date(end, "end") if end else first
        if last < first:
            raise ValueError("The last day is before the first day.")
        return {
            "start": {"dateTime": f"{first.isoformat()}T00:00:00", "timeZone": tz},
            "end": {"dateTime": f"{(last + dt.timedelta(days=1)).isoformat()}T00:00:00", "timeZone": tz},
        }
    if not end:
        raise ValueError("A timed event needs an `end`.")
    s, e = _parse_datetime(start, "start"), _parse_datetime(end, "end")
    if e <= s:
        raise ValueError("The end is not after the start.")
    return {
        "start": {"dateTime": s.replace(tzinfo=None).isoformat(timespec="seconds"), "timeZone": tz},
        "end": {"dateTime": e.replace(tzinfo=None).isoformat(timespec="seconds"), "timeZone": tz},
    }


def _check_show_as(show_as: str) -> str:
    match = next((s for s in SHOW_AS if s.lower() == show_as.strip().lower()), None)
    if match is None:
        raise ValueError(f"`show_as` must be one of: {', '.join(SHOW_AS)}.")
    return match


def _body(text: str) -> dict:
    return {"contentType": "HTML", "content": "<p>" + html.escape(text).replace("\n", "<br>") + "</p>"}


async def list_events(
    graph: Graph, start_date: str, end_date: str, query: str | None = None, limit: int = 100
) -> list[dict]:
    """Events (recurring ones expanded) from the start of `start_date` to the end of `end_date`."""
    first, last = _parse_date(start_date, "start_date"), _parse_date(end_date, "end_date")
    if last < first:
        raise ValueError("`end_date` is before `start_date`.")
    tz = await mailbox_time_zone(graph)
    params = {
        "startDateTime": f"{first.isoformat()}T00:00:00",
        "endDateTime": f"{(last + dt.timedelta(days=1)).isoformat()}T00:00:00",
        "$select": _EVENT_FIELDS,
        "$orderby": "start/dateTime",
        "$top": "100",
    }
    headers = {"Prefer": f'outlook.timezone="{tz}"'}
    # Without a timezone on the window, Graph reads it as UTC; ask for it in the mailbox's zone instead.
    events = await graph.get_all("/me/calendarView", params=params, limit=1000 if query else limit, headers=headers)
    if query:
        q = query.lower()
        events = [e for e in events if q in (e.get("subject") or "").lower()
                  or q in ((e.get("location") or {}).get("displayName") or "").lower()
                  or q in (e.get("bodyPreview") or "").lower()]
    return [summarise(e) for e in events[:limit]]


async def create_event(
    graph: Graph, subject: str, start: str, end: str | None = None, all_day: bool = False,
    show_as: str = "busy", location: str | None = None, notes: str | None = None,
    reminder_minutes: int | None = None,
) -> dict:
    tz = await mailbox_time_zone(graph)
    payload: dict = {"subject": subject, "isAllDay": all_day, "showAs": _check_show_as(show_as),
                     **_times(start, end, all_day, tz)}
    if location:
        payload["location"] = {"displayName": location}
    if notes:
        payload["body"] = _body(notes)
    if reminder_minutes is not None:
        payload["isReminderOn"] = reminder_minutes >= 0
        if reminder_minutes >= 0:
            payload["reminderMinutesBeforeStart"] = reminder_minutes
    event = await graph.post("/me/events", json=payload)
    return summarise(event)


async def update_event(
    graph: Graph, event_id: str, subject: str | None = None, start: str | None = None, end: str | None = None,
    all_day: bool | None = None, show_as: str | None = None, location: str | None = None,
    notes: str | None = None,
) -> dict:
    current = await graph.get(f"/me/events/{event_id}", params={"$select": _EVENT_FIELDS})
    if current.get("attendees"):
        raise ValueError("This event has attendees, and changing it would send them updates. Change it in Outlook.")
    payload: dict = {}
    if subject is not None:
        payload["subject"] = subject
    if show_as is not None:
        payload["showAs"] = _check_show_as(show_as)
    if location is not None:
        payload["location"] = {"displayName": location}
    if notes is not None:
        payload["body"] = _body(notes)
    if start is not None or end is not None or all_day is not None:
        # The times and the all-day flag have to agree, so work out both from what's
        # asked for plus what the event already has.
        was = summarise(current)
        new_all_day = was["all_day"] if all_day is None else all_day
        if all_day is not None and all_day != was["all_day"] and start is None:
            raise ValueError("Give a `start` (and `end`) when switching an event between all-day and timed.")
        new_start = start if start is not None else was["start"]
        new_end = end if end is not None else (None if start is not None and new_all_day else was["end"])
        tz = await mailbox_time_zone(graph)
        payload.update({"isAllDay": new_all_day, **_times(new_start, new_end, new_all_day, tz)})
    if not payload:
        raise ValueError("Nothing to change.")
    event = await graph.patch(f"/me/events/{event_id}", json=payload)
    return summarise(event)
