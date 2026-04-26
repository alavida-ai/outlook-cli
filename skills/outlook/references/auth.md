# Auth

Microsoft device-code flow. No passwords pass through the agent — the user signs in on their own device.

## Synchronous login (the only mode)

```bash
outlook auth login
```

What happens, step by step:

1. CLI hits Microsoft's `/devicecode` endpoint, gets a short URL + 6-character user code (~200 ms)
2. **Prints URL + code to stderr**, then blocks polling Microsoft every ~5 seconds for up to 15 minutes
3. User opens URL on any device (phone is fine), enters the code, signs in with their normal credentials (MFA if enforced)
4. Subprocess unblocks, caches tokens to OS keychain (or 0600 file fallback), exits 0

## Agent pattern (when enrolling a new user)

Spawn `outlook auth login` as a subprocess and read stderr in real time. Stderr is line-buffered so the URL appears immediately on a piped subprocess — no buffering gymnastics needed.

```python
import subprocess

proc = subprocess.Popen(
    ["outlook", "auth", "login"],
    stderr=subprocess.PIPE, stdout=subprocess.DEVNULL, text=True,
)

# The first line of stderr always carries the URL + code.
url_line = proc.stderr.readline()
# url_line: "To sign in, use a web browser to open ... and enter the code ABCD1234 ..."

# Forward to the user via whatever channel the agent has (Slack/SMS/email/...).
send_to_user_via_preferred_channel(url_line)

# Block until the user signs in (typically ~2 minutes once they click the URL).
returncode = proc.wait()
if returncode != 0:
    # User didn't sign in within 15 min, or another error.
    handle_failure()
```

After the subprocess returns 0, tokens are cached. All subsequent `outlook ...` commands silently use them.

## Detecting auth state

```bash
outlook auth status                # exits 0 if cached account exists, 1 if not
outlook whoami                     # exits 1 with a friendly stderr message if not authed
outlook whoami --json              # exits 1 the same way; stdout stays empty so jq is safe
```

Any other Graph-touching command (`outlook mail list`, `outlook calendar create`, etc.) will exit 1 with the same `Not logged in. Run \`outlook auth login\` first.` message if tokens are missing or refresh has failed.

## Re-auth triggers

Tokens silently refresh forever, except in these cases:

- **Password change** → all refresh tokens invalidated server-side; user must re-login
- **Admin consent revocation** → same
- **90-day continuous idle** → refresh token expires
- **Conditional Access re-evaluation** → may force re-login
- **User signs out at https://myaccount.microsoft.com** → tokens invalidated

These are uncommon. Default assumption: once authed, stays authed.

## Logout

```bash
outlook auth logout
```

Clears tokens from OS keychain + file fallback. The next command will require a fresh `outlook auth login`.

## Token storage

| Platform | Backend | Notes |
| --- | --- | --- |
| macOS | Keychain | Native, no setup |
| Linux desktop | Secret Service (libsecret) | Requires GNOME Keyring or KDE Wallet |
| Windows | Credential Manager | Native, no setup |
| Headless Linux (e.g. VPS) | File at `~/.outlook-cli/tokens.json` | 0600 perms, parent dir 0700 |

The CLI tries the OS keychain first; if none is available, it falls back to the encrypted-at-rest file. The CLI never asks for or stores Microsoft passwords — only access + refresh tokens issued by Microsoft via OAuth.

## Multiple accounts

Not yet supported. The CLI manages one cached account per host user. To switch, `logout` + `login` with the new account.
