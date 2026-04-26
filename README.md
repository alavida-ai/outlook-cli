# outlook-cli

Alavida's CLI for Microsoft Outlook — mail, calendar, contacts via Microsoft Graph, designed for AI agents.

Python. Delegated permissions only (agent acts **as** the user). Device-code auth. OS-keychain-backed token storage. Draft-only mail (no send) — human stays in the loop.

## Status

Phase 0 infrastructure + Phase 1 mail (partial) live. Calendar + contacts + remaining mail features in [PLAN.md](PLAN.md). Agent-facing reference in [SKILL.md](SKILL.md).

## Architecture

```
scripts/provision_entra_app.py   # one-shot: register the shared multi-tenant Entra app
src/outlook_cli/
  cli.py                         # Typer root app (`outlook ...`)
  auth.py                        # MSAL device-code flow + keyring token storage
  graph.py                       # msgraph-sdk client factory (auto-refresh)
  commands/
    _common.py                   # shared config helpers + JSON envelope
    auth.py                      # `outlook auth`
    mail.py                      # `outlook mail`
    calendar.py                  # `outlook calendar` (stub)
    contacts.py                  # `outlook contacts` (stub)
```

## Permissions (delegated)

- `Mail.ReadWrite` — read + create/update drafts. **No send.**
- `Calendars.ReadWrite`
- `Calendars.ReadWrite.Shared`
- `Contacts.ReadWrite`
- `User.Read`
- `offline_access` (refresh tokens — added automatically)

Draft-only mail is deliberate: agent prepares, human sends. Matches FCA human-in-the-loop for regulated communications.

## How auth works

The CLI ships with a **shared multi-tenant Entra app** embedded in `auth.py:DEFAULT_CLIENT_ID`. End users never see a client ID or tenant ID. The `common` authority means sign-in works for personal Microsoft accounts and any work/school tenant.

Per-client onboarding is **one consent click** by the client's IT admin:
```
https://login.microsoftonline.com/<their-tenant>/adminconsent?client_id=<our-app-id>
```

End users authenticate once via device code, then it's silent refresh forever.

Tokens live in the OS credential manager (macOS Keychain, Linux Secret Service, Windows Credential Manager), with a 0600-locked file fallback at `~/.outlook-cli/tokens.json` for headless systems.

### Escape hatch: dedicated client app

A paranoid client can run their own Entra app. Override via `.env`:

```
AZURE_CLIENT_ID=<their-app-id>
AZURE_TENANT_ID=<their-tenant-id>
```

## Setup

### First-time (one-off, per outlook-cli install)

1. `uv sync`
2. (Only if bootstrapping a new Alavida app — skip if `DEFAULT_CLIENT_ID` already baked in)
   ```bash
   uv run python scripts/provision_entra_app.py --tenant alavidai.onmicrosoft.com --multi-tenant
   ```
   Paste the printed client id into `src/outlook_cli/auth.py:DEFAULT_CLIENT_ID`.

### Per end user

```bash
uv run outlook auth login
```

Follows the device-code flow — prints a short URL + code to stderr, open URL in any browser, sign in, the command unblocks and caches the tokens.

**For agents:** spawn as a subprocess, read stderr in real time to get the URL + code, forward to the user, wait for the subprocess to exit. Stdout is left clean for machine-readable output; stderr is line-buffered so the URL appears immediately on a pipe. See [SKILL.md](SKILL.md) for a Python example.

## Quick test

```bash
uv run outlook whoami
uv run outlook mail list --limit 5
uv run outlook mail list --unread --json | jq '.count'
```

## OpenClaw skill

The CLI ships with a bundled OpenClaw skill that teaches the agent when and how to use `outlook ...`. Install it onto disk where OpenClaw scans for skills:

```bash
outlook skill install            # → ~/.openclaw/skills/outlook
outlook skill install --force    # overwrite existing
outlook skill uninstall          # remove
outlook skill path               # show bundled source location (read-only)
```

After install, restart OpenClaw to pick it up:
```bash
openclaw gateway restart
openclaw skills list             # should now show 'outlook'
```

The skill source is at `skills/outlook/` in this repo: `SKILL.md` (frontmatter + index) plus `references/{auth,mail,calendar,safety}.md`. They're bundled into the wheel via Hatch's `force-include` and shipped with every install.

## Roadmap

See [PLAN.md](PLAN.md) for the phased plan and [SKILL.md](SKILL.md) for the agent-facing command reference.

## Why not olkcli?

[`rlrghb/olkcli`](https://github.com/rlrghb/olkcli) is the mature reference implementation — well-designed Go CLI wrapping Graph. We're building our own because:

1. Python matches the rest of the Alavida stack
2. Anonymous author + young repo isn't defensible for FCA-regulated clients
3. Draft-only mail is a design constraint, not an option
4. `msgraph-sdk-python` does most of the heavy lifting — small surface to own

olkcli remains a goldmine for CLI shape, scope choices, and agent-focused ergonomics.
