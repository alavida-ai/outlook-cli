"""`outlook auth` sub-app — login, logout, status, two-phase device flow."""

from __future__ import annotations

import json
import sys
from typing import Annotated

import typer

from outlook_cli import auth
from outlook_cli.commands._common import client_id, console, err_console, tenant_id

app = typer.Typer(help="Authentication.", no_args_is_help=True)


@app.command("login")
def login():
    """Synchronous device-code login. Prints the URL + code, blocks until you sign in."""
    auth.login(tenant_id(), client_id())
    err_console.print("[green]Logged in. Token cached.[/green]")


@app.command("login-start")
def login_start(
    as_json: Annotated[bool, typer.Option("--json", help="Emit JSON only.")] = False,
):
    """Begin device flow and return the URL + code for a human to visit.

    Pair with `outlook auth login-complete --handle <h>` to finish after the user signs in.
    Designed for agents: forward the URL/code to the user (Slack, email, SMS, ...), then
    run `login-complete` to poll until sign-in succeeds.
    """
    info = auth.start_device_flow(tenant_id(), client_id())

    if as_json:
        json.dump(info, sys.stdout)
        sys.stdout.write("\n")
        return

    console.print(f"[bold]Verification URL:[/bold] {info['verification_uri']}")
    console.print(f"[bold]Code:[/bold]            {info['user_code']}")
    console.print(f"[dim]Expires in {info['expires_in']}s.[/dim]")
    console.print()
    console.print(f"Finish with: [cyan]outlook auth login-complete --handle {info['handle']}[/cyan]")


@app.command("login-complete")
def login_complete(
    handle: Annotated[str, typer.Option("--handle", help="Handle from `login-start`.")],
    as_json: Annotated[bool, typer.Option("--json", help="Emit JSON only.")] = False,
):
    """Poll until the user completes sign-in. Caches tokens on success."""
    result = auth.complete_device_flow(handle)

    if as_json:
        json.dump({"ok": True, "scopes": result.get("scope", "").split()}, sys.stdout)
        sys.stdout.write("\n")
        return

    err_console.print("[green]Logged in. Token cached.[/green]")


@app.command("logout")
def logout():
    """Delete the cached token."""
    auth.logout()
    err_console.print("[yellow]Logged out.[/yellow]")


@app.command("status")
def status():
    """Show the currently authenticated account."""
    account = auth.status(tenant_id(), client_id())
    if account is None:
        err_console.print("[yellow]Not logged in.[/yellow] Run `outlook auth login`.")
        raise typer.Exit(1)
    console.print_json(json.dumps(account))
