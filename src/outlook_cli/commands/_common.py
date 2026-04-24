"""Shared helpers for command modules: consoles, config resolution, output helpers."""

from __future__ import annotations

import json
import os
import sys
from typing import Any

import typer
from rich.console import Console

from outlook_cli import auth

console = Console()
err_console = Console(stderr=True)


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
