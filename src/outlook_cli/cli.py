"""Typer root app — wires sub-apps from outlook_cli.commands."""

from __future__ import annotations

import json
import sys
from typing import Annotated

import typer
from dotenv import load_dotenv

from outlook_cli import graph
from outlook_cli.commands import auth, calendar, contacts, mail
from outlook_cli.commands._common import client_id, console, run_graph, tenant_id

load_dotenv()

app = typer.Typer(
    name="outlook",
    help="Alavida Outlook CLI — Graph-backed mail, calendar, contacts.",
    no_args_is_help=True,
)
app.add_typer(auth.app, name="auth")
app.add_typer(mail.app, name="mail")
app.add_typer(calendar.app, name="calendar")
app.add_typer(contacts.app, name="contacts")


@app.command("whoami")
def whoami(
    as_json: Annotated[bool, typer.Option("--json", help="Emit JSON.")] = False,
):
    """Show the currently authenticated user's profile."""
    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        return await client.me.get()

    me = run_graph(_run())

    info = {
        "displayName": me.display_name,
        "mail": me.mail or me.user_principal_name,
        "userPrincipalName": me.user_principal_name,
        "jobTitle": me.job_title,
        "department": me.department,
        "officeLocation": me.office_location,
        "id": me.id,
    }

    if as_json:
        json.dump(info, sys.stdout)
        sys.stdout.write("\n")
        return

    for k, v in info.items():
        console.print(f"[cyan]{k:20}[/cyan] {v or ''}")


if __name__ == "__main__":
    app()
