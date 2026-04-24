"""`outlook mail` sub-app — read, search, draft, reply, move, delete."""

from __future__ import annotations

import asyncio
import sys
from pathlib import Path
from typing import Annotated, Any

import typer
from rich.table import Table

from outlook_cli import graph
from outlook_cli.commands._common import (
    client_id,
    console,
    err_console,
    parse_select,
    print_json_envelope,
    tenant_id,
)

app = typer.Typer(help="Mail — read, draft, triage.", no_args_is_help=True)


# Well-known folder aliases Graph recognises natively.
WELL_KNOWN_FOLDERS = {
    "inbox", "sentitems", "drafts", "deleteditems", "junkemail", "archive",
    "outbox", "scheduled", "clutter",
}


def _message_summary(m: Any) -> dict:
    """Flatten a Graph Message into a dict for JSON output."""
    return {
        "id": m.id,
        "subject": m.subject,
        "from": m.from_.email_address.address if m.from_ and m.from_.email_address else None,
        "received": m.received_date_time.isoformat() if m.received_date_time else None,
        "is_read": m.is_read,
        "has_attachments": m.has_attachments,
        "preview": m.body_preview,
        "web_link": m.web_link,
    }


def _message_full(m: Any) -> dict:
    """Extended shape for `mail read`."""
    summary = _message_summary(m)
    summary.update({
        "to": [r.email_address.address for r in (m.to_recipients or []) if r.email_address],
        "cc": [r.email_address.address for r in (m.cc_recipients or []) if r.email_address],
        "bcc": [r.email_address.address for r in (m.bcc_recipients or []) if r.email_address],
        "body": m.body.content if m.body else None,
        "body_content_type": m.body.content_type.value if m.body and m.body.content_type else None,
        "importance": m.importance.value if m.importance else None,
    })
    return summary


# ── list ──────────────────────────────────────────────────────────────────

@app.command("list")
def list_(
    limit: Annotated[int, typer.Option("-n", "--limit", help="Max messages.")] = 10,
    folder: Annotated[str, typer.Option("-f", "--folder", help="Folder name (inbox, sentitems, drafts, ...) or id.")] = "inbox",
    unread: Annotated[bool, typer.Option("-u", "--unread", help="Only unread.")] = False,
    from_addr: Annotated[str | None, typer.Option("--from", help="Filter by sender address.")] = None,
    as_json: Annotated[bool, typer.Option("--json", help="JSON envelope.")] = False,
    select: Annotated[str | None, typer.Option("--select", help="Comma-separated field projection.")] = None,
):
    """List messages in a folder (default: inbox)."""
    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
            MessagesRequestBuilder,
        )

        filters = []
        if unread:
            filters.append("isRead eq false")
        if from_addr:
            filters.append(f"from/emailAddress/address eq '{from_addr}'")
        filter_expr = " and ".join(filters) if filters else None

        qp = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=limit,
            orderby=["receivedDateTime DESC"],
            select=["id", "subject", "from", "receivedDateTime", "isRead", "hasAttachments", "bodyPreview", "webLink"],
            filter=filter_expr,
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=qp,
        )

        folder_key = folder.lower()
        if folder_key in WELL_KNOWN_FOLDERS or _looks_like_id(folder):
            page = await client.me.mail_folders.by_mail_folder_id(folder_key if folder_key in WELL_KNOWN_FOLDERS else folder).messages.get(request_configuration=config)
        else:
            # Treat as a custom folder display name — resolve to id.
            folder_id = await _resolve_folder_id(client, folder)
            page = await client.me.mail_folders.by_mail_folder_id(folder_id).messages.get(request_configuration=config)

        return page.value or []

    messages = asyncio.run(_run())

    if as_json:
        print_json_envelope([_message_summary(m) for m in messages], fields=parse_select(select))
        return

    table = Table(title=f"{folder} (top {len(messages)})")
    table.add_column("Received", style="cyan", no_wrap=True)
    table.add_column("From", style="magenta")
    table.add_column("Subject")
    table.add_column("📎", justify="center")
    table.add_column("•", justify="center")
    for m in messages:
        table.add_row(
            m.received_date_time.strftime("%Y-%m-%d %H:%M") if m.received_date_time else "",
            m.from_.email_address.address if m.from_ and m.from_.email_address else "",
            m.subject or "",
            "•" if m.has_attachments else "",
            " " if m.is_read else "●",
        )
    console.print(table)


# ── read ──────────────────────────────────────────────────────────────────

@app.command("read")
def read(
    message_id: Annotated[str, typer.Argument(help="Message id (from `mail list`).")],
    as_json: Annotated[bool, typer.Option("--json", help="Emit full JSON.")] = False,
    prefer_text: Annotated[bool, typer.Option("--text", help="Request plain-text body (default: HTML).")] = False,
):
    """Read a single message in full."""
    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.users.item.messages.item.message_item_request_builder import (
            MessageItemRequestBuilder,
        )

        config = MessageItemRequestBuilder.MessageItemRequestBuilderGetRequestConfiguration()
        if prefer_text:
            config.headers.add("Prefer", "outlook.body-content-type=text")
        return await client.me.messages.by_message_id(message_id).get(request_configuration=config)

    msg = asyncio.run(_run())

    if as_json:
        import json as _json
        _json.dump(_message_full(msg), sys.stdout, default=str)
        sys.stdout.write("\n")
        return

    console.rule(msg.subject or "(no subject)")
    console.print(f"[cyan]From:[/cyan]    {msg.from_.email_address.address if msg.from_ and msg.from_.email_address else ''}")
    console.print(f"[cyan]To:[/cyan]      {', '.join(r.email_address.address for r in (msg.to_recipients or []) if r.email_address)}")
    if msg.cc_recipients:
        console.print(f"[cyan]Cc:[/cyan]      {', '.join(r.email_address.address for r in msg.cc_recipients if r.email_address)}")
    console.print(f"[cyan]Date:[/cyan]    {msg.received_date_time}")
    console.print()
    if msg.body:
        console.print(msg.body.content or "")


# ── search ────────────────────────────────────────────────────────────────

@app.command("search")
def search(
    query: Annotated[str, typer.Argument(help="KQL query (e.g. `from:boss@co.com subject:urgent`).")],
    limit: Annotated[int, typer.Option("-n", "--limit", help="Max results.")] = 25,
    as_json: Annotated[bool, typer.Option("--json", help="JSON envelope.")] = False,
    select: Annotated[str | None, typer.Option("--select", help="Comma-separated field projection.")] = None,
):
    """Search across all mail folders using KQL."""
    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.users.item.messages.messages_request_builder import (
            MessagesRequestBuilder,
        )

        qp = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=limit,
            search=f'"{query}"',
            select=["id", "subject", "from", "receivedDateTime", "isRead", "hasAttachments", "bodyPreview", "webLink"],
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=qp,
        )
        page = await client.me.messages.get(request_configuration=config)
        return page.value or []

    messages = asyncio.run(_run())

    if as_json:
        print_json_envelope([_message_summary(m) for m in messages], fields=parse_select(select))
        return

    table = Table(title=f'Search: "{query}"  ({len(messages)} results)')
    table.add_column("Received", style="cyan", no_wrap=True)
    table.add_column("From", style="magenta")
    table.add_column("Subject")
    for m in messages:
        table.add_row(
            m.received_date_time.strftime("%Y-%m-%d %H:%M") if m.received_date_time else "",
            m.from_.email_address.address if m.from_ and m.from_.email_address else "",
            m.subject or "",
        )
    console.print(table)


# ── draft ─────────────────────────────────────────────────────────────────

@app.command("draft")
def draft(
    to: Annotated[list[str], typer.Option("--to", help="Recipient address (repeat for multiple).")],
    subject: Annotated[str, typer.Option("--subject", help="Email subject.")],
    body: Annotated[str | None, typer.Option("--body", help="Body text. Use '-' to read from stdin.")] = None,
    body_file: Annotated[Path | None, typer.Option("--body-file", help="Read body from this file.")] = None,
    cc: Annotated[list[str] | None, typer.Option("--cc", help="CC address (repeatable).")] = None,
    bcc: Annotated[list[str] | None, typer.Option("--bcc", help="BCC address (repeatable).")] = None,
    html: Annotated[bool, typer.Option("--html", help="Treat body as HTML instead of plain text.")] = False,
    as_json: Annotated[bool, typer.Option("--json", help="Emit raw JSON.")] = False,
):
    """Create a draft in the Drafts folder. Does not send."""
    body_text = _resolve_body(body, body_file)

    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.email_address import EmailAddress
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.message import Message
        from msgraph.generated.models.recipient import Recipient

        def _r(addrs: list[str] | None) -> list[Recipient]:
            return [Recipient(email_address=EmailAddress(address=a)) for a in (addrs or [])]

        msg = Message(
            subject=subject,
            body=ItemBody(
                content_type=BodyType.Html if html else BodyType.Text,
                content=body_text,
            ),
            to_recipients=_r(to),
            cc_recipients=_r(cc),
            bcc_recipients=_r(bcc),
        )
        return await client.me.messages.post(msg)

    created = asyncio.run(_run())

    if as_json:
        import json as _json
        _json.dump({"id": created.id, "subject": created.subject, "web_link": created.web_link}, sys.stdout)
        sys.stdout.write("\n")
        return

    err_console.print(f"[green]Draft created.[/green] id={created.id}")
    if created.web_link:
        err_console.print(f"Open in Outlook: {created.web_link}")


# ── reply / reply-all / forward ───────────────────────────────────────────

@app.command("reply")
def reply(
    message_id: Annotated[str, typer.Argument(help="Message id to reply to.")],
    body: Annotated[str | None, typer.Option("--body", help="Reply body. Use '-' for stdin.")] = None,
    body_file: Annotated[Path | None, typer.Option("--body-file", help="Read body from file.")] = None,
    reply_all: Annotated[bool, typer.Option("--all", help="Reply to everyone on the thread.")] = False,
    html: Annotated[bool, typer.Option("--html", help="Treat body as HTML.")] = False,
    as_json: Annotated[bool, typer.Option("--json", help="Emit JSON.")] = False,
):
    """Create a draft reply (or reply-all). Does not send."""
    body_text = _resolve_body(body, body_file)

    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.message import Message
        from msgraph.generated.users.item.messages.item.create_reply_all.create_reply_all_post_request_body import (
            CreateReplyAllPostRequestBody,
        )
        from msgraph.generated.users.item.messages.item.create_reply.create_reply_post_request_body import (
            CreateReplyPostRequestBody,
        )

        reply_msg = Message(body=ItemBody(
            content_type=BodyType.Html if html else BodyType.Text,
            content=body_text,
        ))

        builder = client.me.messages.by_message_id(message_id)
        if reply_all:
            body_req = CreateReplyAllPostRequestBody(message=reply_msg)
            return await builder.create_reply_all.post(body_req)
        body_req = CreateReplyPostRequestBody(message=reply_msg)
        return await builder.create_reply.post(body_req)

    draft_msg = asyncio.run(_run())

    if as_json:
        import json as _json
        _json.dump({"id": draft_msg.id, "web_link": draft_msg.web_link}, sys.stdout)
        sys.stdout.write("\n")
        return

    err_console.print(f"[green]Draft reply created.[/green] id={draft_msg.id}")
    if draft_msg.web_link:
        err_console.print(f"Open in Outlook: {draft_msg.web_link}")


@app.command("forward")
def forward(
    message_id: Annotated[str, typer.Argument(help="Message id to forward.")],
    to: Annotated[list[str], typer.Option("--to", help="Recipient (repeatable).")],
    comment: Annotated[str | None, typer.Option("--comment", help="Optional leading note.")] = None,
    as_json: Annotated[bool, typer.Option("--json", help="Emit JSON.")] = False,
):
    """Create a draft forward. Does not send."""
    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.models.email_address import EmailAddress
        from msgraph.generated.models.recipient import Recipient
        from msgraph.generated.users.item.messages.item.create_forward.create_forward_post_request_body import (
            CreateForwardPostRequestBody,
        )

        body = CreateForwardPostRequestBody(
            comment=comment or "",
            to_recipients=[Recipient(email_address=EmailAddress(address=a)) for a in to],
        )
        return await client.me.messages.by_message_id(message_id).create_forward.post(body)

    draft_msg = asyncio.run(_run())

    if as_json:
        import json as _json
        _json.dump({"id": draft_msg.id, "web_link": draft_msg.web_link}, sys.stdout)
        sys.stdout.write("\n")
        return

    err_console.print(f"[green]Draft forward created.[/green] id={draft_msg.id}")


# ── move / delete / mark ──────────────────────────────────────────────────

@app.command("move")
def move(
    message_id: Annotated[str, typer.Argument()],
    folder: Annotated[str, typer.Argument(help="Target folder (well-known name or id).")],
):
    """Move a message to another folder."""
    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.users.item.messages.item.move.move_post_request_body import (
            MovePostRequestBody,
        )

        folder_key = folder.lower()
        if folder_key in WELL_KNOWN_FOLDERS or _looks_like_id(folder):
            dest_id = folder_key if folder_key in WELL_KNOWN_FOLDERS else folder
        else:
            dest_id = await _resolve_folder_id(client, folder)

        body = MovePostRequestBody(destination_id=dest_id)
        return await client.me.messages.by_message_id(message_id).move.post(body)

    asyncio.run(_run())
    err_console.print(f"[green]Moved[/green] {message_id} -> {folder}")


@app.command("delete")
def delete(
    message_id: Annotated[str, typer.Argument()],
    force: Annotated[bool, typer.Option("--force", help="Skip confirmation.")] = False,
):
    """Delete a message (moves to Deleted Items)."""
    if not force:
        typer.confirm(f"Delete message {message_id}?", abort=True)
    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        await client.me.messages.by_message_id(message_id).delete()

    asyncio.run(_run())
    err_console.print(f"[green]Deleted[/green] {message_id}")


@app.command("mark")
def mark(
    message_id: Annotated[str, typer.Argument()],
    read_flag: Annotated[bool, typer.Option("--read/--unread", help="Mark read or unread.")] = True,
):
    """Mark a message read or unread."""
    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.models.message import Message
        patch = Message(is_read=read_flag)
        await client.me.messages.by_message_id(message_id).patch(patch)

    asyncio.run(_run())
    state = "read" if read_flag else "unread"
    err_console.print(f"[green]Marked {state}[/green] {message_id}")


# ── folders ───────────────────────────────────────────────────────────────

@app.command("folders")
def folders(
    as_json: Annotated[bool, typer.Option("--json", help="JSON envelope.")] = False,
):
    """List mail folders."""
    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        page = await client.me.mail_folders.get()
        return page.value or []

    fs = asyncio.run(_run())

    if as_json:
        print_json_envelope([
            {
                "id": f.id,
                "displayName": f.display_name,
                "unreadItemCount": f.unread_item_count,
                "totalItemCount": f.total_item_count,
            }
            for f in fs
        ])
        return

    table = Table(title="Mail folders")
    table.add_column("Name")
    table.add_column("Unread", justify="right", style="cyan")
    table.add_column("Total", justify="right")
    table.add_column("ID", style="dim")
    for f in fs:
        table.add_row(f.display_name, str(f.unread_item_count or 0), str(f.total_item_count or 0), f.id)
    console.print(table)


# ── helpers ───────────────────────────────────────────────────────────────

def _resolve_body(body: str | None, body_file: Path | None) -> str:
    if body is not None and body_file is not None:
        err_console.print("[red]--body and --body-file are mutually exclusive.[/red]")
        raise typer.Exit(2)
    if body == "-":
        return sys.stdin.read()
    if body is not None:
        return body
    if body_file is not None:
        return body_file.read_text()
    # No body provided — read stdin if it's piped.
    if not sys.stdin.isatty():
        return sys.stdin.read()
    err_console.print("[red]Provide --body, --body-file, or pipe via stdin.[/red]")
    raise typer.Exit(2)


def _looks_like_id(s: str) -> bool:
    # Graph folder/message ids are long base64-ish strings. Heuristic: >20 chars, not all lowercase letters.
    return len(s) > 20


async def _resolve_folder_id(client, display_name: str) -> str:
    """Look up a custom folder's id by its displayName."""
    from msgraph.generated.users.item.mail_folders.mail_folders_request_builder import (
        MailFoldersRequestBuilder,
    )

    qp = MailFoldersRequestBuilder.MailFoldersRequestBuilderGetQueryParameters(
        filter=f"displayName eq '{display_name}'",
    )
    config = MailFoldersRequestBuilder.MailFoldersRequestBuilderGetRequestConfiguration(
        query_parameters=qp,
    )
    page = await client.me.mail_folders.get(request_configuration=config)
    if not page or not page.value:
        raise typer.Exit(f"Folder '{display_name}' not found.")
    return page.value[0].id
