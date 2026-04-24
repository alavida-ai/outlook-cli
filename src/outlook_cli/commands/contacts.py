"""`outlook contacts` sub-app — address book."""

from __future__ import annotations

import typer

from outlook_cli.commands._common import err_console

app = typer.Typer(help="Contacts.", no_args_is_help=True)


@app.command("list")
def list_():
    """List contacts. (stub — TODO)"""
    err_console.print("[yellow]TODO:[/yellow] contacts list.")
