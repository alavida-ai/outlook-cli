"""Shared helpers for command modules: consoles, config resolution, output helpers."""

from __future__ import annotations

import asyncio
import json
import os
import sys
from typing import Any, Awaitable, TypeVar

import typer
from rich.console import Console

from outlook_cli import auth

console = Console()
err_console = Console(stderr=True)

T = TypeVar("T")


def run_graph(coro: Awaitable[T]) -> T:
    """Run an async Graph call, translating common errors into clean CLI output.

    Catches the 'not logged in' / 'silent refresh failed' RuntimeError from auth.py
    and exits 1 with a friendly stderr message instead of dumping a traceback through
    the Graph SDK's request pipeline.
    """
    try:
        return asyncio.run(coro)
    except RuntimeError as e:
        msg = str(e)
        if "No cached account" in msg or "Silent token refresh failed" in msg:
            err_console.print(
                "[yellow]Not logged in.[/yellow] Run `outlook auth login` first."
            )
            raise typer.Exit(1) from None
        raise


def tenant_id() -> str:
    """Resolve tenant id. Env override > embedded default ('common')."""
    return os.environ.get("AZURE_TENANT_ID") or auth.DEFAULT_TENANT


def client_id() -> str:
    """Resolve client id. Env override > embedded default. Errors if neither set."""
    val = os.environ.get("AZURE_CLIENT_ID") or auth.DEFAULT_CLIENT_ID
    if not val:
        err_console.print(
            "[red]No client ID available.[/red]\n"
            "Either set AZURE_CLIENT_ID in .env, or run "
            "`uv run python scripts/provision_entra_app.py --tenant <t> --multi-tenant` "
            "and paste the returned client id into "
            "src/outlook_cli/auth.py:DEFAULT_CLIENT_ID."
        )
        raise typer.Exit(1)
    return val


def print_json_envelope(results: list[Any], next_link: str | None = None, fields: list[str] | None = None) -> None:
    """Print a JSON envelope matching the olkcli shape: {results, count, nextLink}.

    If `fields` is provided, project each result dict to only those keys.
    """
    if fields:
        results = [{k: r.get(k) for k in fields} for r in results]
    envelope = {
        "results": results,
        "count": len(results),
        "nextLink": next_link,
    }
    json.dump(envelope, sys.stdout, default=str)
    sys.stdout.write("\n")


def parse_select(select: str | None) -> list[str] | None:
    """Parse `--select from,subject,id` into a list."""
    if not select:
        return None
    return [s.strip() for s in select.split(",") if s.strip()]


def interpret_escapes(s: str) -> str:
    r"""Interpret backslash escapes in a free-text arg the way printf does.

    Common case we're protecting against: an agent calls
        outlook ... --body "Hi\n\nfoo"
    Bash double-quoting does not interpret backslash escapes, so the CLI
    receives the literal sequence backslash + n. Without this, that goes
    straight to Graph and the rendered email/event/comment shows the
    backslashes verbatim.

    Decodes \\n, \\r, \\t, \\\\ only. Real newlines and tabs already in
    the input pass through unchanged because we only react to a literal
    backslash followed by one of those escape characters.
    """
    out = []
    i = 0
    while i < len(s):
        if s[i] == "\\" and i + 1 < len(s):
            nxt = s[i + 1]
            if nxt == "n":
                out.append("\n")
                i += 2
                continue
            if nxt == "r":
                out.append("\r")
                i += 2
                continue
            if nxt == "t":
                out.append("\t")
                i += 2
                continue
            if nxt == "\\":
                out.append("\\")
                i += 2
                continue
        out.append(s[i])
        i += 1
    return "".join(out)
