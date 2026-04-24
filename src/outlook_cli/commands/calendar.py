"""`outlook calendar` sub-app — events, scheduling, availability."""

from __future__ import annotations

import json
import sys
from datetime import datetime, timedelta, timezone
from typing import Annotated, Any

import typer
from rich.table import Table

from outlook_cli import graph
from outlook_cli.commands._common import (
    client_id,
    console,
    err_console,
    parse_select,
    print_json_envelope,
    run_graph,
    tenant_id,
)

app = typer.Typer(help="Calendar events.", no_args_is_help=True)

RECURRENCE_PRESETS = {"daily", "weekdays", "weekly", "monthly", "yearly"}
ATTENDEE_RESPONSES = {"accept", "decline", "tentative"}


def _iso8601(s: str) -> str:
    """Normalise a user-supplied date/datetime into the form Graph accepts.

    Accepts: '2026-04-15', '2026-04-15T09:00', '2026-04-15T09:00:00', '...Z'.
    Returns an ISO-8601 string without trailing Z (Graph takes tz separately).
    """
    s = s.rstrip("Z").rstrip("z")
    if "T" not in s:
        s += "T00:00:00"
    elif s.count(":") == 1:
        s += ":00"
    return s


def _event_summary(e: Any) -> dict:
    """Flatten Graph Event for JSON output."""
    return {
        "id": e.id,
        "subject": e.subject,
        "start": e.start.date_time if e.start else None,
        "end": e.end.date_time if e.end else None,
        "time_zone": e.start.time_zone if e.start else None,
        "location": e.location.display_name if e.location else None,
        "organizer": (
            e.organizer.email_address.address
            if e.organizer and e.organizer.email_address else None
        ),
        "attendees": [
            {
                "address": a.email_address.address if a.email_address else None,
                "name": a.email_address.name if a.email_address else None,
                "type": a.type.value if a.type else None,
                "response": (
                    a.status.response.value if a.status and a.status.response else None
                ),
            }
            for a in (e.attendees or [])
        ],
        "is_online_meeting": e.is_online_meeting,
        "online_join_url": (
            e.online_meeting.join_url if e.online_meeting else None
        ),
        "is_all_day": e.is_all_day,
        "is_cancelled": e.is_cancelled,
        "web_link": e.web_link,
    }


# ── list ──────────────────────────────────────────────────────────────────

@app.command("list")
def list_(
    days: Annotated[int, typer.Option("-d", "--days", help="How many days forward to show.")] = 7,
    after: Annotated[str | None, typer.Option("--after", help="Start datetime (overrides --days).")] = None,
    before: Annotated[str | None, typer.Option("--before", help="End datetime (overrides --days).")] = None,
    limit: Annotated[int, typer.Option("-n", "--limit", help="Max events.")] = 50,
    as_json: Annotated[bool, typer.Option("--json", help="JSON envelope.")] = False,
    select: Annotated[str | None, typer.Option("--select", help="Comma-separated fields.")] = None,
):
    """List events in a date range. Uses calendarView — recurring events are expanded."""
    now = datetime.now(timezone.utc).replace(microsecond=0)
    after_iso = _iso8601(after) if after else now.isoformat().split("+")[0]
    before_iso = (
        _iso8601(before) if before
        else (now + timedelta(days=days)).isoformat().split("+")[0]
    )

    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.users.item.calendar_view.calendar_view_request_builder import (
            CalendarViewRequestBuilder,
        )
        qp = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters(
            start_date_time=after_iso,
            end_date_time=before_iso,
            top=limit,
            orderby=["start/dateTime"],
        )
        config = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetRequestConfiguration(
            query_parameters=qp,
        )
        page = await client.me.calendar_view.get(request_configuration=config)
        return page.value or []

    events = run_graph(_run())

    if as_json:
        print_json_envelope([_event_summary(e) for e in events], fields=parse_select(select))
        return

    table = Table(title=f"Events {after_iso[:16]} → {before_iso[:16]}  ({len(events)})")
    table.add_column("Start", style="cyan", no_wrap=True)
    table.add_column("Subject")
    table.add_column("Attendees", style="dim")
    table.add_column("Location", style="dim")
    for e in events:
        start_str = e.start.date_time[:16].replace("T", " ") if e.start else ""
        attendees = ", ".join(
            a.email_address.address for a in (e.attendees or []) if a.email_address
        )
        if len(attendees) > 40:
            attendees = attendees[:37] + "..."
        loc = e.location.display_name if e.location else ""
        table.add_row(start_str, e.subject or "", attendees, loc)
    console.print(table)


# ── show ──────────────────────────────────────────────────────────────────

@app.command("show")
def show(
    event_id: Annotated[str, typer.Argument(help="Event id.")],
    as_json: Annotated[bool, typer.Option("--json", help="Emit full JSON.")] = False,
):
    """Show a single event in full."""
    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        return await client.me.events.by_event_id(event_id).get()

    e = run_graph(_run())

    if as_json:
        json.dump(_event_summary(e), sys.stdout, default=str)
        sys.stdout.write("\n")
        return

    console.rule(e.subject or "(no subject)")
    start = e.start.date_time if e.start else ""
    end = e.end.date_time if e.end else ""
    tz_ = e.start.time_zone if e.start else ""
    console.print(f"[cyan]When:[/cyan]      {start} → {end} ({tz_})")
    if e.location and e.location.display_name:
        console.print(f"[cyan]Where:[/cyan]     {e.location.display_name}")
    if e.organizer and e.organizer.email_address:
        console.print(f"[cyan]Organizer:[/cyan] {e.organizer.email_address.address}")
    if e.attendees:
        console.print("[cyan]Attendees:[/cyan]")
        for a in e.attendees:
            addr = a.email_address.address if a.email_address else ""
            resp = a.status.response.value if a.status and a.status.response else ""
            console.print(f"  → {addr}  [dim]({resp})[/dim]")
    if e.online_meeting and e.online_meeting.join_url:
        console.print(f"[cyan]Join:[/cyan]      {e.online_meeting.join_url}")
    console.print()
    if e.body and e.body.content:
        console.print(e.body.content)
    if e.web_link:
        console.print()
        console.print(f"[dim]Open in Outlook: {e.web_link}[/dim]")


# ── create ────────────────────────────────────────────────────────────────

@app.command("create")
def create(
    subject: Annotated[str, typer.Option("--subject", help="Event title.")],
    start: Annotated[str, typer.Option("--start", help="Start ISO 8601 (e.g. 2026-04-15T09:00).")],
    end: Annotated[str, typer.Option("--end", help="End ISO 8601.")],
    attendees: Annotated[list[str] | None, typer.Option("--attendees", help="Attendee email (repeatable).")] = None,
    location: Annotated[str | None, typer.Option("--location", help="Location display name.")] = None,
    body: Annotated[str | None, typer.Option("--body", help="Event body / description.")] = None,
    all_day: Annotated[bool, typer.Option("--all-day", help="All-day event.")] = False,
    online_meeting: Annotated[bool, typer.Option("--online-meeting", help="Add a Teams meeting link.")] = False,
    recurrence: Annotated[str | None, typer.Option("--recurrence", help="daily | weekdays | weekly | monthly | yearly")] = None,
    tz: Annotated[str, typer.Option("--tz", help="IANA timezone for start/end.")] = "UTC",
    as_json: Annotated[bool, typer.Option("--json", help="Emit JSON.")] = False,
):
    """Create a calendar event. Sends invites to attendees."""
    if recurrence and recurrence not in RECURRENCE_PRESETS:
        err_console.print(
            f"[red]--recurrence must be one of: {', '.join(sorted(RECURRENCE_PRESETS))}[/red]"
        )
        raise typer.Exit(2)

    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.models.attendee import Attendee
        from msgraph.generated.models.attendee_type import AttendeeType
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
        from msgraph.generated.models.day_of_week import DayOfWeek
        from msgraph.generated.models.email_address import EmailAddress
        from msgraph.generated.models.event import Event
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.location import Location
        from msgraph.generated.models.online_meeting_provider_type import OnlineMeetingProviderType
        from msgraph.generated.models.patterned_recurrence import PatternedRecurrence
        from msgraph.generated.models.recurrence_pattern import RecurrencePattern
        from msgraph.generated.models.recurrence_pattern_type import RecurrencePatternType
        from msgraph.generated.models.recurrence_range import RecurrenceRange
        from msgraph.generated.models.recurrence_range_type import RecurrenceRangeType

        event = Event(
            subject=subject,
            start=DateTimeTimeZone(date_time=_iso8601(start), time_zone=tz),
            end=DateTimeTimeZone(date_time=_iso8601(end), time_zone=tz),
            is_all_day=all_day,
        )
        if body is not None:
            event.body = ItemBody(content_type=BodyType.Text, content=body)
        if location:
            event.location = Location(display_name=location)
        if attendees:
            event.attendees = [
                Attendee(email_address=EmailAddress(address=a), type=AttendeeType.Required)
                for a in attendees
            ]
        if online_meeting:
            event.is_online_meeting = True
            event.online_meeting_provider = OnlineMeetingProviderType.TeamsForBusiness
        if recurrence:
            patterns = {
                "daily": RecurrencePattern(
                    type=RecurrencePatternType.Daily, interval=1
                ),
                "weekdays": RecurrencePattern(
                    type=RecurrencePatternType.Weekly,
                    interval=1,
                    days_of_week=[
                        DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday,
                        DayOfWeek.Thursday, DayOfWeek.Friday,
                    ],
                ),
                "weekly": RecurrencePattern(
                    type=RecurrencePatternType.Weekly,
                    interval=1,
                    days_of_week=[_iso_weekday_enum(_iso8601(start))],
                ),
                "monthly": RecurrencePattern(
                    type=RecurrencePatternType.AbsoluteMonthly,
                    interval=1,
                    day_of_month=int(_iso8601(start)[8:10]),
                ),
                "yearly": RecurrencePattern(
                    type=RecurrencePatternType.AbsoluteYearly,
                    interval=1,
                    day_of_month=int(_iso8601(start)[8:10]),
                    month=int(_iso8601(start)[5:7]),
                ),
            }
            event.recurrence = PatternedRecurrence(
                pattern=patterns[recurrence],
                range=RecurrenceRange(
                    type=RecurrenceRangeType.NoEnd,
                    start_date=_iso8601(start)[:10],
                ),
            )

        return await client.me.events.post(event)

    created = run_graph(_run())

    if as_json:
        json.dump(_event_summary(created), sys.stdout, default=str)
        sys.stdout.write("\n")
        return

    err_console.print(f"[green]Event created.[/green] id={created.id}")
    if created.online_meeting and created.online_meeting.join_url:
        err_console.print(f"Join link: {created.online_meeting.join_url}")
    if created.web_link:
        err_console.print(f"Open in Outlook: {created.web_link}")


def _iso_weekday_enum(iso_datetime: str):
    """Map ISO date → msgraph DayOfWeek enum."""
    from msgraph.generated.models.day_of_week import DayOfWeek

    days = [
        DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday,
        DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday,
    ]
    dt = datetime.fromisoformat(iso_datetime)
    return days[dt.weekday()]


# ── update ────────────────────────────────────────────────────────────────

@app.command("update")
def update(
    event_id: Annotated[str, typer.Argument()],
    subject: Annotated[str | None, typer.Option("--subject")] = None,
    start: Annotated[str | None, typer.Option("--start")] = None,
    end: Annotated[str | None, typer.Option("--end")] = None,
    location: Annotated[str | None, typer.Option("--location")] = None,
    body: Annotated[str | None, typer.Option("--body")] = None,
    tz: Annotated[str, typer.Option("--tz")] = "UTC",
):
    """Update an existing event. Only the fields you pass get PATCHed."""
    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
        from msgraph.generated.models.event import Event
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.location import Location

        patch = Event()
        if subject is not None:
            patch.subject = subject
        if start is not None:
            patch.start = DateTimeTimeZone(date_time=_iso8601(start), time_zone=tz)
        if end is not None:
            patch.end = DateTimeTimeZone(date_time=_iso8601(end), time_zone=tz)
        if location is not None:
            patch.location = Location(display_name=location)
        if body is not None:
            patch.body = ItemBody(content_type=BodyType.Text, content=body)

        return await client.me.events.by_event_id(event_id).patch(patch)

    run_graph(_run())
    err_console.print(f"[green]Updated[/green] {event_id}")


# ── delete ────────────────────────────────────────────────────────────────

@app.command("delete")
def delete(
    event_id: Annotated[str, typer.Argument()],
    force: Annotated[bool, typer.Option("--force", help="Skip confirmation.")] = False,
):
    """Delete an event. Notifies attendees if any."""
    if not force:
        typer.confirm(f"Delete event {event_id}? This may notify attendees.", abort=True)

    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        await client.me.events.by_event_id(event_id).delete()

    run_graph(_run())
    err_console.print(f"[green]Deleted[/green] {event_id}")


# ── respond ──────────────────────────────────────────────────────────────

@app.command("respond")
def respond(
    event_id: Annotated[str, typer.Argument()],
    response: Annotated[str, typer.Argument(help="accept | decline | tentative")],
    comment: Annotated[str | None, typer.Option("--comment", help="Optional note to organizer.")] = None,
    send_response: Annotated[bool, typer.Option("--send/--no-send", help="Notify organizer.")] = True,
):
    """Accept, decline, or tentatively accept an incoming meeting invite."""
    if response not in ATTENDEE_RESPONSES:
        err_console.print("[red]response must be: accept | decline | tentative[/red]")
        raise typer.Exit(2)

    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        if response == "accept":
            from msgraph.generated.users.item.events.item.accept.accept_post_request_body import (
                AcceptPostRequestBody,
            )
            body = AcceptPostRequestBody(send_response=send_response, comment=comment)
            await client.me.events.by_event_id(event_id).accept.post(body)
        elif response == "decline":
            from msgraph.generated.users.item.events.item.decline.decline_post_request_body import (
                DeclinePostRequestBody,
            )
            body = DeclinePostRequestBody(send_response=send_response, comment=comment)
            await client.me.events.by_event_id(event_id).decline.post(body)
        else:
            from msgraph.generated.users.item.events.item.tentatively_accept.tentatively_accept_post_request_body import (
                TentativelyAcceptPostRequestBody,
            )
            body = TentativelyAcceptPostRequestBody(send_response=send_response, comment=comment)
            await client.me.events.by_event_id(event_id).tentatively_accept.post(body)

    run_graph(_run())
    err_console.print(f"[green]Responded {response}[/green] to {event_id}")


# ── availability (free-busy) ──────────────────────────────────────────────

@app.command("availability")
def availability(
    emails: Annotated[list[str], typer.Option("--emails", help="Email to check (repeatable).")],
    days: Annotated[int, typer.Option("-d", "--days", help="How many days forward.")] = 7,
    interval: Annotated[int, typer.Option("--interval", help="Availability view interval in minutes.")] = 60,
    tz: Annotated[str, typer.Option("--tz", help="IANA timezone for the query window.")] = "UTC",
    as_json: Annotated[bool, typer.Option("--json")] = False,
):
    """Check free/busy across one or more users (including yourself).

    Availability view legend: 0=free 1=tentative 2=busy 3=out-of-office 4=working-elsewhere.
    """
    start = datetime.now(timezone.utc).replace(microsecond=0)
    end = start + timedelta(days=days)

    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
        from msgraph.generated.users.item.calendar.get_schedule.get_schedule_post_request_body import (
            GetSchedulePostRequestBody,
        )

        body = GetSchedulePostRequestBody(
            schedules=emails,
            start_time=DateTimeTimeZone(
                date_time=start.isoformat().split("+")[0], time_zone=tz,
            ),
            end_time=DateTimeTimeZone(
                date_time=end.isoformat().split("+")[0], time_zone=tz,
            ),
            availability_view_interval=interval,
        )
        resp = await client.me.calendar.get_schedule.post(body)
        return resp.value or []

    schedules = run_graph(_run())

    if as_json:
        data = [
            {
                "email": s.schedule_id,
                "availability_view": s.availability_view,
                "items": [
                    {
                        "subject": item.subject,
                        "start": item.start.date_time if item.start else None,
                        "end": item.end.date_time if item.end else None,
                        "status": item.status.value if item.status else None,
                    }
                    for item in (s.schedule_items or [])
                ],
            }
            for s in schedules
        ]
        print_json_envelope(data)
        return

    table = Table(
        title=f"Availability  window={days}d  block={interval}min  "
              "0=free 1=tentative 2=busy 3=OOO 4=elsewhere",
    )
    table.add_column("Email")
    table.add_column("Availability view", style="cyan")
    for s in schedules:
        table.add_row(s.schedule_id or "", s.availability_view or "")
    console.print(table)
