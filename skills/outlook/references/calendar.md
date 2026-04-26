# Calendar

Events with attendees send invites the moment they're created — there is no calendar-draft mode equivalent to mail drafts. Confirm attendees + timing with the user before running `calendar create --attendees ...`. See safety.md for the confirmation rules.

## List

```bash
outlook calendar list                                  # next 7 days
outlook calendar list -d 30                            # next 30 days
outlook calendar list --after 2026-05-01 --before 2026-05-08
outlook calendar list --json
outlook calendar list --json --select subject,start,end
```

Uses `calendarView` under the hood, which **expands recurring events** to individual occurrences. Each row in the result is a single occurrence, even for recurring meetings.

JSON shape per event:
```json
{
  "id": "...",
  "subject": "...",
  "start": "2026-05-01T10:00:00.0000000",
  "end": "2026-05-01T11:00:00.0000000",
  "time_zone": "UTC",
  "location": "Conference Room A",
  "organizer": "alice@example.com",
  "attendees": [
    {"address": "...", "name": "...", "type": "required", "response": "accepted"}
  ],
  "is_online_meeting": true,
  "online_join_url": "https://teams.microsoft.com/...",
  "is_all_day": false,
  "is_cancelled": false,
  "web_link": "https://outlook.office365.com/owa/?itemid=..."
}
```

## Show one event

```bash
outlook calendar show <event-id>
outlook calendar show <event-id> --json
```

Renders subject, when, where, organizer, attendees + their responses, join URL, body.

## Create

```bash
outlook calendar create \
  --subject "Quick sync" \
  --start 2026-05-01T14:00 \
  --end 2026-05-01T14:30 \
  --tz "Europe/London"
```

```bash
# With attendees — sends invites immediately
outlook calendar create --subject "Project kickoff" \
  --start 2026-05-01T10:00 --end 2026-05-01T11:00 \
  --attendees alice@example.com --attendees bob@example.com \
  --location "Conference Room A" \
  --body "Agenda: ..."
```

```bash
# Teams online meeting
outlook calendar create --subject "Remote sync" \
  --start 2026-05-01T15:00 --end 2026-05-01T15:30 \
  --online-meeting
# Result includes online_join_url (Teams meeting link)
```

```bash
# All-day event
outlook calendar create --subject "Offsite" \
  --start 2026-06-01 --end 2026-06-02 --all-day
```

```bash
# Recurring (presets: daily | weekdays | weekly | monthly | yearly)
outlook calendar create --subject "Weekly standup" \
  --start 2026-05-04T09:00 --end 2026-05-04T09:15 \
  --recurrence weekdays
```

Time zones: pass an IANA name with `--tz` (default `UTC`). Same tz applied to start and end.

Output: stderr shows `Event created. id=...`, the Teams join URL if `--online-meeting`, and an `Open in Outlook: <web_link>` line.

## Update (PATCH — only the fields you pass)

```bash
outlook calendar update <id> --subject "New subject"
outlook calendar update <id> --start 2026-05-01T15:00 --end 2026-05-01T15:30
outlook calendar update <id> --location "Room B"
outlook calendar update <id> --body "Updated agenda"
```

If the event has attendees, updating start/end will trigger an invite-update notification to all of them.

## Delete

```bash
outlook calendar delete <id>           # interactive confirmation
outlook calendar delete <id> --force   # skip prompt — NOTIFIES ATTENDEES, get user OK first
```

Deletion of an event with attendees fires a cancellation notification to all of them. Real-world consequence — confirm before passing `--force`.

## Respond to incoming invites

```bash
outlook calendar respond <id> accept
outlook calendar respond <id> decline --comment "Conflict — let's reschedule"
outlook calendar respond <id> tentative --no-send   # don't notify organizer
```

Default behaviour notifies the organizer of your response. `--no-send` suppresses that notification (the organizer's view of your status updates next time they check the event).

## Free/busy across users

```bash
outlook calendar availability \
  --emails a@b.com --emails c@d.com \
  -d 7 --interval 60
```

```bash
# Higher resolution (30-min blocks), 5 days, JSON output
outlook calendar availability \
  --emails alice@example.com --emails bob@example.com \
  -d 5 --interval 30 --json
```

Returns a compact "availability view" string per user — each character is one block:

| Char | Meaning |
| --- | --- |
| `0` | Free |
| `1` | Tentative |
| `2` | Busy |
| `3` | Out of office |
| `4` | Working elsewhere |

E.g. `"0000022200002220000000"` means free for the first 5 blocks, busy for 3, free for 4, busy for 3, free again. With a 30-min interval and 1-day window, you get 48 chars per user.

JSON output also includes a per-user `items` array with `subject`, `start`, `end`, `status` for actual conflicts (when the responding user has chosen to share details).

## Common workflows

### "Am I free at 3pm tomorrow?"

```bash
outlook calendar list --after 2026-05-02T15:00 --before 2026-05-02T16:00 --json | jq '.count'
# 0 → free; >0 → check the events.
```

### "Find a 30-min slot for me + Alex tomorrow afternoon"

```bash
outlook calendar availability \
  --emails me@example.com --emails alex@example.com \
  -d 1 --interval 30 --json
# Look for runs of consecutive '0's across both users' availability_view strings.
```

### "What meetings do I have this week?"

```bash
outlook calendar list -d 7 --json | jq -r '.results[] | "\(.start)  \(.subject)"'
```

### "Cancel my 3pm with Alex" (confirm first!)

```bash
ID=$(outlook calendar list --after 2026-05-02T15:00 --before 2026-05-02T16:00 --json | jq -r '.results[0].id')
# Surface to user: "About to delete event 'X' at Y with Z attendees — confirm?"
# After they confirm:
outlook calendar delete "$ID" --force
```
