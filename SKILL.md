---
name: outlook
description: Microsoft Outlook CLI — read mail, draft messages, triage inbox as the user via Microsoft Graph. Draft-only mail (no send). Delegated permissions only.
homepage: https://github.com/alavida-ai/outlook-cli
metadata: {"openclaw":{"emoji":"📬","homepage":"https://github.com/alavida-ai/outlook-cli","os":["darwin","linux"],"requires":{"bins":["outlook"]},"install":[{"id":"uv","kind":"uv","package":"git+https://github.com/alavida-ai/outlook-cli","bins":["outlook"],"label":"Install outlook-cli (uv)"}]}}
---

# outlook — Microsoft Outlook for AI agents

Use `outlook` to read mail, draft messages, search, and triage the user's inbox **as them** via delegated Microsoft Graph permissions. All mail writes are drafts — the human reviews and sends from Outlook.

## Install

```bash
uv tool install git+https://github.com/alavida-ai/outlook-cli
```

Then `outlook --help`. If you don't have `uv`: `curl -LsSf https://astral.sh/uv/install.sh | sh` first.

## Authenticate (first time, per user)

The CLI uses Microsoft's **device-code flow** — no passwords pass through your machine. The user visits a short URL on any device, enters a 6-character code, signs in with their normal credentials (MFA if enforced).

### Agent-mediated login (the normal path)

When an agent is enrolling a user, split login into two calls so the agent can relay the link:

```bash
# 1. Start the flow — emits a JSON line with URL + code, exits immediately
outlook auth login-start --json
# → {"verification_uri":"https://microsoft.com/devicelogin","user_code":"AB12CD34","handle":"abc123...","expires_in":900,"message":"..."}

# 2. Forward verification_uri + user_code to the user via whatever channel
#    (Slack, email, SMS). They open the URL, enter the code, sign in.

# 3. Finish — polls Microsoft until the user completes sign-in, caches tokens
outlook auth login-complete --handle abc123...
```

The `handle` from step 1 is an opaque token that lets step 3 resume the same flow. Flow state is persisted to `$TMPDIR/outlook-cli-flows/<handle>.json` (0600). Two different processes can own the two calls — step 1 from the agent's dispatch loop, step 3 from a waiter.

### Synchronous login (manual)

If a human is at a terminal:

```bash
outlook auth login
# Prints: "To sign in, visit https://microsoft.com/devicelogin and enter code AB12CD34"
# Blocks for up to 15 min until the user completes sign-in
```

### Re-auth triggers

Once signed in, tokens silently refresh forever. Re-login is only needed on:
- Password change
- Admin consent revocation
- 90-day idle
- Conditional Access re-evaluation events

If any Graph call returns an auth error, run `outlook auth login-start --json` and relay the URL/code again.

### Other auth commands

- `outlook auth status` — show the cached account (exits 1 if not logged in)
- `outlook auth logout` — clear all cached tokens
- `outlook whoami [--json]` — display name, email, job title, department, office

## Commands

### Mail — read

- `outlook mail list [-n 10] [-f inbox|sentitems|drafts|archive|<name>|<id>] [-u] [--from addr] [--json] [--select fields]`
  - `-f inbox` is default. Well-known folders: `inbox`, `sentitems`, `drafts`, `deleteditems`, `junkemail`, `archive`. Or pass a custom folder name or folder id.
  - `-u` shows only unread.
  - `--from someone@domain.com` filters by sender.
- `outlook mail read <id> [--text] [--json]`
  - `--text` requests a plain-text body via `Prefer: outlook.body-content-type=text` (default is HTML — noisier for LLMs).
- `outlook mail search "<kql>" [-n 25] [--json]`
  - KQL examples: `from:boss@co.com`, `subject:invoice`, `hasattachment:true`, `received>=2026-04-01`.
- `outlook mail folders [--json]` — list folders with unread/total counts.

### Mail — write (all drafts, never sends)

- `outlook mail draft --to X --subject Y (--body Z | --body-file path | -) [--cc X] [--bcc X] [--html] [--json]`
  - Use `--body -` to read the body from stdin.
  - Creates a draft in the user's Drafts folder. The user reviews in Outlook and sends.
- `outlook mail reply <id> --body Z [--all] [--html] [--json]` — draft reply (or reply-all).
- `outlook mail forward <id> --to X [--comment "..."]` — draft forward.

### Mail — triage

- `outlook mail move <id> <folder>` — well-known name or id.
- `outlook mail delete <id> [--force]` — moves to Deleted Items.
- `outlook mail mark <id> [--read|--unread]` — toggle read state.

### Calendar, contacts (not yet implemented)

- `outlook calendar list [--days N]` — stub, returns TODO.
- `outlook contacts list` — stub, returns TODO.

## Output

Every command that returns data supports `--json`. List commands use an envelope:

```json
{ "results": [...], "count": 10, "nextLink": null }
```

Single-item commands (`mail read`, `mail draft`, `mail reply`, `mail forward`, `whoami`, `auth login-start`, `auth login-complete`) emit a bare object.

`--select field1,field2` projects each result to only those keys — reduces token cost when you only need a subset.

Human text goes to **stderr** (status, errors). Data goes to **stdout**. Safe to pipe stdout to `jq` without contamination.

## Design constraints (why it works the way it does)

- **No `mail send`.** Every mail write is a draft. Human reviews and sends from Outlook. Strong compliance story for regulated clients (FCA, MiFID II).
- **Delegated-only.** The CLI never holds application permissions (tenant-wide access). It acts as the signed-in user; every Graph call is scoped to that user's mailbox.
- **No audit layer inside the CLI.** Outlook's own folders (Drafts, Sent Items, calendar) + the M365 Purview Unified Audit Log already capture everything. Agent-level auditing (which prompt, which tool call) belongs upstream in the agent framework.

## Scopes requested at login

`Mail.ReadWrite`, `Calendars.ReadWrite`, `Calendars.ReadWrite.Shared`, `Contacts.ReadWrite`, `User.Read`, `offline_access`.

## Error handling

Non-zero exit codes on failure. Human error text on stderr. When `--json` is set, data stays on stdout, errors on stderr — always safe to `| jq`.

## Agent scripting examples

```bash
# Count unread in inbox
outlook mail list --unread --json | jq '.count'

# Subjects from a specific sender
outlook mail list --from boss@co.com --json | jq -r '.results[].subject'

# Draft from stdin
echo "Status update..." | outlook mail draft --to team@example.com --subject "Friday update" --body -

# Search then read the first hit as plain text
ID=$(outlook mail search "subject:invoice" --json | jq -r '.results[0].id')
outlook mail read "$ID" --text

# Triage pattern: move everything from a sender to a folder
outlook mail list --from notifications@vendor.com --json \
  | jq -r '.results[].id' \
  | xargs -I{} outlook mail move {} archive
```

## Override the embedded app (paranoid clients only)

The CLI ships with a shared multi-tenant Entra app id baked in. For a client that wants their own dedicated app in their own tenant, set:

```
AZURE_CLIENT_ID=<their-app-id>
AZURE_TENANT_ID=<their-tenant-id>
```

in a `.env` file in the working directory. The env vars override the embedded defaults.
