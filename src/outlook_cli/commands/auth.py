"""`outlook auth` sub-app — login, logout, status."""

from __future__ import annotations

import json

import typer

from outlook_cli import auth
from outlook_cli.commands._common import client_id, console, err_console, tenant_id

app = typer.Typer(help="Authentication.", no_args_is_help=True)


@app.command("login")
def login():
    """Device-code login. Prints the URL + code to stderr, blocks until you sign in."""
    auth.login(tenant_id(), client_id())
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
