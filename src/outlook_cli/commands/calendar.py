"""`outlook calendar` sub-app — events on the user's calendar."""

from __future__ import annotations

from typing import Annotated

import typer

from outlook_cli.commands._common import err_console

app = typer.Typer(help="Calendar events.", no_args_is_help=True)


@app.command("list")
def list_(
    days: Annotated[int, typer.Option(help="How many days forward to show.")] = 7,
):
    """List upcoming calendar events. (stub — TODO)"""
    err_console.print(f"[yellow]TODO:[/yellow] list events for next {days} days.")
