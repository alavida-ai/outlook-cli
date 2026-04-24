"""Device-code authentication + MSAL token cache.

Two sources of config, in priority order:

1. Environment variables `AZURE_CLIENT_ID` / `AZURE_TENANT_ID` (escape hatch — e.g. a
   client that wants their own dedicated Entra app).
2. Embedded defaults below — the normal path, ships with the package.

Tokens are stored in the OS credential manager (macOS Keychain, Linux Secret Service,
Windows Credential Manager) via `keyring`. On systems without a backend (e.g. headless
Linux VPS without libsecret), falls back to a 0600-locked file at
`~/.outlook-cli/tokens.json`.
"""

from __future__ import annotations

import json
import os
import tempfile
import time
from pathlib import Path

import keyring
import msal
from keyring.errors import KeyringError

# ─────────────────────────────────────────────────────────────────────────────
# Embedded app identity. After running `provision_entra_app.py --multi-tenant`,
# paste the resulting app_id here. Leave as None to require AZURE_CLIENT_ID env.
# ─────────────────────────────────────────────────────────────────────────────
DEFAULT_CLIENT_ID: str | None = "18f9e6ff-2b0a-423e-bb35-ab9b541e604e"

# 'common' works for both personal Microsoft accounts and any work/school tenant.
# MSAL resolves the actual tenant from the user's sign-in. Multi-tenant Entra apps
# require this (or 'organizations' to exclude personal accounts).
DEFAULT_TENANT: str = "common"

SCOPES = [
    "Mail.ReadWrite",
    "Calendars.ReadWrite",
    "Calendars.ReadWrite.Shared",
    "Contacts.ReadWrite",
    "User.Read",
]
# `offline_access` is added automatically when a public client requests scopes.

KEYRING_SERVICE = "outlook-cli"
KEYRING_KEY = "default"


def _cache_path() -> Path:
    override = os.environ.get("OUTLOOK_CLI_TOKEN_CACHE")
    if override:
        return Path(override).expanduser()
    return Path.home() / ".outlook-cli" / "tokens.json"


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    try:
        data = keyring.get_password(KEYRING_SERVICE, KEYRING_KEY)
        if data:
            cache.deserialize(data)
            return cache
    except KeyringError:
        pass
    path = _cache_path()
    if path.exists():
        cache.deserialize(path.read_text())
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if not cache.has_state_changed:
        return
    blob = cache.serialize()
    try:
        keyring.set_password(KEYRING_SERVICE, KEYRING_KEY, blob)
        return
    except KeyringError:
        pass
    path = _cache_path()
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    path.write_text(blob)
    path.chmod(0o600)


def _build_app(tenant_id: str, client_id: str, cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
    )


def login(tenant_id: str, client_id: str) -> dict:
    """Run the device-code flow synchronously. Prints the URL + code, then blocks on poll.

    For an agent that wants to forward the login link to a human and come back later,
    use `start_device_flow` + `complete_device_flow` instead.
    """
    cache = _load_cache()
    app = _build_app(tenant_id, client_id, cache)

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to start device flow: {json.dumps(flow)}")

    print(flow["message"])  # "To sign in, visit https://microsoft.com/devicelogin and enter code XXXX"

    result = app.acquire_token_by_device_flow(flow)
    _save_cache(cache)

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")
    return result


# ── Two-phase device flow (for agents) ────────────────────────────────────


def _flow_dir() -> Path:
    return Path(tempfile.gettempdir()) / "outlook-cli-flows"


def _flow_path(handle: str) -> Path:
    safe = "".join(c for c in handle if c.isalnum() or c in "-_")
    return _flow_dir() / f"{safe}.json"


def start_device_flow(tenant_id: str, client_id: str) -> dict:
    """Begin device-code auth and return the code + URL without blocking.

    The returned dict contains:
        verification_uri — where the user should sign in
        user_code        — the short code they type there
        expires_in       — seconds before the code expires (typically 900)
        handle           — opaque token to pass to `complete_device_flow`
        message          — human-readable instructions

    The MSAL flow state is persisted to a temp file keyed by `handle` so a later
    invocation can pick it up (different process / later in the agent's lifecycle).
    """
    cache = _load_cache()
    app = _build_app(tenant_id, client_id, cache)

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to start device flow: {json.dumps(flow)}")

    handle = flow["device_code"][:16]
    _flow_dir().mkdir(parents=True, exist_ok=True)
    path = _flow_path(handle)
    path.write_text(json.dumps({
        "flow": flow,
        "tenant_id": tenant_id,
        "client_id": client_id,
    }))
    path.chmod(0o600)

    return {
        "verification_uri": flow["verification_uri"],
        "user_code": flow["user_code"],
        "expires_in": flow.get("expires_in", 900),
        "handle": handle,
        "message": flow["message"],
    }


def complete_device_flow(handle: str) -> dict:
    """Poll Microsoft until the user completes sign-in. Caches tokens on success.

    Call after `start_device_flow` once the user has been given the URL + code.
    Blocks for up to `expires_in` seconds (default 15 min).
    """
    path = _flow_path(handle)
    if not path.exists():
        raise RuntimeError(f"Unknown flow handle: {handle}. Did `start_device_flow` run?")

    state = json.loads(path.read_text())
    flow = state["flow"]
    tenant_id = state["tenant_id"]
    client_id = state["client_id"]

    cache = _load_cache()
    app = _build_app(tenant_id, client_id, cache)
    result = app.acquire_token_by_device_flow(flow)
    _save_cache(cache)

    try:
        path.unlink()
    except OSError:
        pass

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")
    return result


def get_access_token(tenant_id: str, client_id: str) -> tuple[str, int]:
    """Return (access_token, expires_on_unix_ts). Refreshes silently if needed."""
    cache = _load_cache()
    app = _build_app(tenant_id, client_id, cache)

    accounts = app.get_accounts()
    if not accounts:
        raise RuntimeError("No cached account. Run `outlook auth login` first.")

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    _save_cache(cache)

    if not result or "access_token" not in result:
        raise RuntimeError(
            "Silent token refresh failed. Run `outlook auth login` to re-authenticate."
        )

    expires_on = int(time.time()) + int(result.get("expires_in", 3600))
    return result["access_token"], expires_on


def logout() -> None:
    """Delete the token cache from keychain and file fallback."""
    try:
        keyring.delete_password(KEYRING_SERVICE, KEYRING_KEY)
    except KeyringError:
        pass
    path = _cache_path()
    if path.exists():
        path.unlink()


def status(tenant_id: str, client_id: str) -> dict | None:
    """Return info about the cached account, or None if not logged in."""
    cache = _load_cache()
    app = _build_app(tenant_id, client_id, cache)
    accounts = app.get_accounts()
    if not accounts:
        return None
    return accounts[0]
