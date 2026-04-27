"""`outlook mail` sub-app — read, search, draft, reply, move, delete."""

from __future__ import annotations

import base64
import binascii
import mimetypes
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path
from typing import Annotated, Any
from urllib.parse import parse_qs, quote, urlparse

import typer
from rich.table import Table

from outlook_cli import graph
from outlook_cli.commands._common import (
    client_id,
    console,
    err_console,
    interpret_escapes,
    parse_select,
    print_json_envelope,
    run_graph,
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


def _compose_link(message: Any) -> str | None:
    """Build an outlook.cloud.microsoft compose URL for a draft/reply/forward.

    Opens the item directly in edit mode while keeping the full Outlook UI
    (folder sidebar, ribbon, reading pane context) visible. Unlike the read-mode
    `webLink` Graph returns for drafts, this saves the user the Edit-pencil click.

    Extracts the URL-encoded ItemID from Graph's webLink — that's the exact
    form the `/mail/compose/<id>` route accepts.
    """
    if not message or not getattr(message, "web_link", None):
        return None
    qs = parse_qs(urlparse(message.web_link).query)
    item_id = qs.get("ItemID", [None])[0]
    if not item_id:
        return None
    return f"https://outlook.cloud.microsoft/mail/compose/{quote(item_id, safe='')}"


# ── list ──────────────────────────────────────────────────────────────────

@app.command("list")
def list_(
    limit: Annotated[int, typer.Option("-n", "--limit", help="Max messages.")] = 10,
    folder: Annotated[str, typer.Option("-f", "--folder", help="Folder name (inbox, sentitems, drafts, ...) or id.")] = "inbox",
    unread: Annotated[bool, typer.Option("-u", "--unread", help="Only unread.")] = False,
    from_addr: Annotated[str | None, typer.Option("--from", help="Filter by sender address.")] = None,
    after: Annotated[str | None, typer.Option("--after", help="Only messages received on/after this date (YYYY-MM-DD or ISO 8601).")] = None,
    before: Annotated[str | None, typer.Option("--before", help="Only messages received on/before this date (YYYY-MM-DD or ISO 8601).")] = None,
    focused: Annotated[bool, typer.Option("--focused", help="Only Focused Inbox messages.")] = False,
    other: Annotated[bool, typer.Option("--other", help="Only Other (non-Focused) messages.")] = False,
    as_json: Annotated[bool, typer.Option("--json", help="JSON envelope.")] = False,
    select: Annotated[str | None, typer.Option("--select", help="Comma-separated field projection.")] = None,
):
    """List messages in a folder (default: inbox)."""
    if focused and other:
        err_console.print("[red]--focused and --other are mutually exclusive.[/red]")
        raise typer.Exit(2)

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
        if after:
            filters.append(f"receivedDateTime ge {_normalise_date(after)}")
        if before:
            filters.append(f"receivedDateTime le {_normalise_date(before)}")
        if focused:
            filters.append("inferenceClassification eq 'focused'")
        if other:
            filters.append("inferenceClassification eq 'other'")
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

    messages = run_graph(_run())

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

    msg = run_graph(_run())

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
    if msg.web_link:
        console.print()
        console.print(f"[dim]Open in Outlook: {msg.web_link}[/dim]")


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

    messages = run_graph(_run())

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
    body: Annotated[str | None, typer.Option("--body", help="Body text. Interprets \\n, \\r, \\t, \\\\ like printf. Use '-' to read from stdin.")] = None,
    body_file: Annotated[Path | None, typer.Option("--body-file", help="Read body from this file (no escape interpretation).")] = None,
    cc: Annotated[list[str] | None, typer.Option("--cc", help="CC address (repeatable).")] = None,
    bcc: Annotated[list[str] | None, typer.Option("--bcc", help="BCC address (repeatable).")] = None,
    html: Annotated[bool, typer.Option("--html", help="Treat body as HTML instead of plain text.")] = False,
    raw_body: Annotated[bool, typer.Option("--raw-body", help="Disable escape interpretation in --body (pass through verbatim).")] = False,
    as_json: Annotated[bool, typer.Option("--json", help="Emit raw JSON.")] = False,
):
    """Create a draft in the Drafts folder. Does not send."""
    body_text = _resolve_body(body, body_file, raw=raw_body)

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

    created = run_graph(_run())
    edit_link = _compose_link(created)

    if as_json:
        import json as _json
        _json.dump({
            "id": created.id,
            "subject": created.subject,
            "web_link": created.web_link,
            "edit_link": edit_link,
        }, sys.stdout)
        sys.stdout.write("\n")
        return

    err_console.print(f"[green]Draft created.[/green] id={created.id}")
    if edit_link:
        err_console.print(f"Edit in Outlook: {edit_link}")


# ── reply / reply-all / forward ───────────────────────────────────────────

@app.command("reply")
def reply(
    message_id: Annotated[str, typer.Argument(help="Message id to reply to.")],
    body: Annotated[str | None, typer.Option("--body", help="Reply body. Interprets \\n, \\r, \\t, \\\\ like printf. Use '-' for stdin.")] = None,
    body_file: Annotated[Path | None, typer.Option("--body-file", help="Read body from file (no escape interpretation).")] = None,
    reply_all: Annotated[bool, typer.Option("--all", help="Reply to everyone on the thread.")] = False,
    html: Annotated[bool, typer.Option("--html", help="Treat body as HTML.")] = False,
    raw_body: Annotated[bool, typer.Option("--raw-body", help="Disable escape interpretation in --body.")] = False,
    as_json: Annotated[bool, typer.Option("--json", help="Emit JSON.")] = False,
):
    """Create a draft reply (or reply-all). Does not send."""
    body_text = _resolve_body(body, body_file, raw=raw_body)

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

    draft_msg = run_graph(_run())
    edit_link = _compose_link(draft_msg)

    if as_json:
        import json as _json
        _json.dump({
            "id": draft_msg.id,
            "web_link": draft_msg.web_link,
            "edit_link": edit_link,
        }, sys.stdout)
        sys.stdout.write("\n")
        return

    err_console.print(f"[green]Draft reply created.[/green] id={draft_msg.id}")
    if edit_link:
        err_console.print(f"Edit in Outlook: {edit_link}")


@app.command("forward")
def forward(
    message_id: Annotated[str, typer.Argument(help="Message id to forward.")],
    to: Annotated[list[str], typer.Option("--to", help="Recipient (repeatable).")],
    comment: Annotated[str | None, typer.Option("--comment", help="Optional leading note. Interprets \\n, \\r, \\t, \\\\ like printf.")] = None,
    raw_comment: Annotated[bool, typer.Option("--raw-comment", help="Disable escape interpretation in --comment.")] = False,
    as_json: Annotated[bool, typer.Option("--json", help="Emit JSON.")] = False,
):
    """Create a draft forward. Does not send."""
    client = graph.get_client(tenant_id(), client_id())

    decoded_comment = comment
    if comment and not raw_comment:
        decoded_comment = interpret_escapes(comment)

    async def _run():
        from msgraph.generated.models.email_address import EmailAddress
        from msgraph.generated.models.recipient import Recipient
        from msgraph.generated.users.item.messages.item.create_forward.create_forward_post_request_body import (
            CreateForwardPostRequestBody,
        )

        body = CreateForwardPostRequestBody(
            comment=decoded_comment or "",
            to_recipients=[Recipient(email_address=EmailAddress(address=a)) for a in to],
        )
        return await client.me.messages.by_message_id(message_id).create_forward.post(body)

    draft_msg = run_graph(_run())
    edit_link = _compose_link(draft_msg)

    if as_json:
        import json as _json
        _json.dump({
            "id": draft_msg.id,
            "web_link": draft_msg.web_link,
            "edit_link": edit_link,
        }, sys.stdout)
        sys.stdout.write("\n")
        return

    err_console.print(f"[green]Draft forward created.[/green] id={draft_msg.id}")
    if edit_link:
        err_console.print(f"Edit in Outlook: {edit_link}")


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

    run_graph(_run())
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

    run_graph(_run())
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

    run_graph(_run())
    state = "read" if read_flag else "unread"
    err_console.print(f"[green]Marked {state}[/green] {message_id}")


# ── flag / importance ─────────────────────────────────────────────────────

@app.command("flag")
def flag(
    message_id: Annotated[str, typer.Argument()],
    status: Annotated[str, typer.Argument(help="flagged | complete | notFlagged")],
):
    """Set the follow-up flag on a message."""
    mapping = {
        "flagged": "flagged",
        "complete": "complete",
        "notflagged": "notFlagged",
    }
    key = status.lower()
    if key not in mapping:
        err_console.print("[red]status must be: flagged | complete | notFlagged[/red]")
        raise typer.Exit(2)

    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.models.followup_flag import FollowupFlag
        from msgraph.generated.models.followup_flag_status import FollowupFlagStatus
        from msgraph.generated.models.message import Message

        status_enum = {
            "flagged": FollowupFlagStatus.Flagged,
            "complete": FollowupFlagStatus.Complete,
            "notflagged": FollowupFlagStatus.NotFlagged,
        }[key]
        patch = Message(flag=FollowupFlag(flag_status=status_enum))
        await client.me.messages.by_message_id(message_id).patch(patch)

    run_graph(_run())
    err_console.print(f"[green]Flag set {mapping[key]}[/green] on {message_id}")


@app.command("importance")
def importance(
    message_id: Annotated[str, typer.Argument()],
    level: Annotated[str, typer.Argument(help="low | normal | high")],
):
    """Set the importance level of a message."""
    key = level.lower()
    if key not in {"low", "normal", "high"}:
        err_console.print("[red]level must be: low | normal | high[/red]")
        raise typer.Exit(2)

    client = graph.get_client(tenant_id(), client_id())

    async def _run():
        from msgraph.generated.models.importance import Importance
        from msgraph.generated.models.message import Message

        imp = {"low": Importance.Low, "normal": Importance.Normal, "high": Importance.High}[key]
        patch = Message(importance=imp)
        await client.me.messages.by_message_id(message_id).patch(patch)

    run_graph(_run())
    err_console.print(f"[green]Importance set {key}[/green] on {message_id}")


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

    fs = run_graph(_run())

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


# ── attachments ───────────────────────────────────────────────────────────

# Graph caps inline FileAttachment POSTs at 3MB. Anything larger needs an upload session.
_INLINE_ATTACHMENT_THRESHOLD = 3 * 1024 * 1024
# Chunk size for resumable uploads. Must be a multiple of 320 KB per Graph guidance.
_UPLOAD_CHUNK_SIZE = 10 * 1024 * 1024
# Hard cap on a single download to prevent runaway memory use.
_MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024
# Trusted suffixes for the pre-authenticated upload URL Graph returns.
_TRUSTED_UPLOAD_HOST_SUFFIXES = (".microsoft.com", ".outlook.com", ".office.com", ".office365.com")
# Ephemeral download root — auto-GC'd on every `attachments` invocation.
_TMP_ROOT = Path.home() / ".outlook-cli" / "tmp"
_TMP_TTL_SECONDS = 24 * 60 * 60

_ODATA_KIND = {
    "#microsoft.graph.fileAttachment": "file",
    "#microsoft.graph.itemAttachment": "item",
    "#microsoft.graph.referenceAttachment": "reference",
}


def _attachment_summary(a: Any) -> dict:
    """Flatten an Attachment (any subclass) into a JSON-friendly dict."""
    return {
        "id": a.id,
        "name": a.name,
        "contentType": a.content_type,
        "size": a.size,
        "isInline": a.is_inline,
        "kind": _ODATA_KIND.get(a.odata_type or "", a.odata_type or "unknown"),
    }


@app.command("attachments")
def attachments(
    message_id: Annotated[str, typer.Argument(help="Message id (from `mail list`).")],
    save: Annotated[bool, typer.Option("--save", help="Download all FileAttachments to --out.")] = False,
    attachment_id: Annotated[str | None, typer.Option("--attachment-id", help="Download a specific attachment by id.")] = None,
    out: Annotated[Path | None, typer.Option("--out", help="Output directory for downloads.")] = None,
    tmp: Annotated[bool, typer.Option("--tmp", help="Download to ~/.outlook-cli/tmp/<msg-id>/ (auto-cleaned after 24h).")] = False,
    as_json: Annotated[bool, typer.Option("--json", help="JSON envelope.")] = False,
):
    """List, or download (--save / --attachment-id), attachments on a message."""
    if save and attachment_id:
        err_console.print("[red]--save and --attachment-id are mutually exclusive.[/red]")
        raise typer.Exit(2)
    if tmp and out is not None:
        err_console.print("[red]--tmp and --out are mutually exclusive.[/red]")
        raise typer.Exit(2)

    # Garbage-collect stale tmp dirs every time we touch attachments.
    _gc_tmp_root()

    if tmp:
        out = _tmp_dir_for_message(message_id)
    elif out is None:
        out = Path(".")

    client = graph.get_client(tenant_id(), client_id())

    # Single-attachment download.
    if attachment_id:
        async def _run_one():
            return await client.me.messages.by_message_id(message_id).attachments.by_attachment_id(attachment_id).get()

        att = run_graph(_run_one())
        saved = _download_one(att, out)
        if as_json:
            import json as _json
            _json.dump({"saved": saved, "name": att.name, "size": att.size}, sys.stdout)
            sys.stdout.write("\n")
            return
        err_console.print(f"[green]Saved[/green] {saved}")
        return

    if save:
        # Single async session: list, then re-fetch each FileAttachment for content_bytes.
        async def _run_bulk():
            page = await client.me.messages.by_message_id(message_id).attachments.get()
            listed = page.value or []
            full_atts: list[Any] = []
            skipped_local: list[dict] = []
            for a in listed:
                kind = _ODATA_KIND.get(a.odata_type or "", "unknown")
                if kind != "file":
                    skipped_local.append({"id": a.id, "name": a.name, "kind": kind})
                    continue
                full = await client.me.messages.by_message_id(message_id).attachments.by_attachment_id(a.id).get()
                full_atts.append(full)
            return full_atts, skipped_local

        full_attachments, skipped = run_graph(_run_bulk())

        _validate_out_dir(out)
        for s in skipped:
            err_console.print(f"[yellow]Skipped[/yellow] {s['name'] or s['id']} (kind={s['kind']})")

        saved_paths: list[str] = []
        for full in full_attachments:
            saved_paths.append(_download_one(full, out, _skip_validate=True))
            err_console.print(f"[green]Saved[/green] {saved_paths[-1]}")

        if as_json:
            import json as _json
            _json.dump({"saved": saved_paths, "skipped": skipped}, sys.stdout)
            sys.stdout.write("\n")
        return

    async def _run_list():
        page = await client.me.messages.by_message_id(message_id).attachments.get()
        return page.value or []

    items = run_graph(_run_list())

    # Listing.
    if as_json:
        print_json_envelope([_attachment_summary(a) for a in items])
        return

    if not items:
        console.print("[dim]No attachments.[/dim]")
        return

    table = Table(title=f"Attachments on {message_id} ({len(items)})")
    table.add_column("Name")
    table.add_column("Type", style="cyan")
    table.add_column("Size", justify="right")
    table.add_column("Kind", style="magenta")
    table.add_column("Inline", justify="center")
    table.add_column("ID", style="dim")
    for a in items:
        kind = _ODATA_KIND.get(a.odata_type or "", "unknown")
        table.add_row(
            a.name or "",
            a.content_type or "",
            _format_bytes(a.size),
            kind,
            "•" if a.is_inline else "",
            a.id or "",
        )
    console.print(table)


@app.command("attach")
def attach(
    draft_id: Annotated[str, typer.Argument(help="Draft message id (from `mail draft`).")],
    file: Annotated[Path, typer.Option("--file", help="File to attach.", exists=True, dir_okay=False, readable=True)],
    name: Annotated[str | None, typer.Option("--name", help="Override the displayed filename.")] = None,
    as_json: Annotated[bool, typer.Option("--json", help="Emit JSON.")] = False,
):
    """Attach a file to an existing draft. Refuses anything other than a draft."""
    file_size = file.stat().st_size
    display_name = name or file.name
    content_type = mimetypes.guess_type(str(file))[0] or "application/octet-stream"

    client = graph.get_client(tenant_id(), client_id())

    # All Graph calls happen inside a single async block — kiota's httpx pool is bound
    # to the first event loop, so multiple run_graph() invocations in one CLI run break.
    async def _run_graph_part():
        from msgraph.generated.users.item.messages.item.message_item_request_builder import (
            MessageItemRequestBuilder,
        )

        # Verify it's a draft up front — Graph's own error otherwise is opaque.
        qp = MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters(
            select=["id", "isDraft"],
        )
        config = MessageItemRequestBuilder.MessageItemRequestBuilderGetRequestConfiguration(
            query_parameters=qp,
        )
        msg = await client.me.messages.by_message_id(draft_id).get(request_configuration=config)
        if not msg.is_draft:
            return ("not_draft", None)

        if file_size <= _INLINE_ATTACHMENT_THRESHOLD:
            from msgraph.generated.models.file_attachment import FileAttachment

            fa = FileAttachment(
                odata_type="#microsoft.graph.fileAttachment",
                name=display_name,
                content_type=content_type,
                content_bytes=file.read_bytes(),
            )
            created = await client.me.messages.by_message_id(draft_id).attachments.post(fa)
            return ("inline", created.id if created else None)

        # Large file — return an upload URL for the sync PUT loop.
        from msgraph.generated.models.attachment_item import AttachmentItem
        from msgraph.generated.models.attachment_type import AttachmentType
        from msgraph.generated.users.item.messages.item.attachments.create_upload_session.create_upload_session_post_request_body import (
            CreateUploadSessionPostRequestBody,
        )

        body = CreateUploadSessionPostRequestBody(
            attachment_item=AttachmentItem(
                attachment_type=AttachmentType.File,
                name=display_name,
                size=file_size,
                content_type=content_type,
            ),
        )
        session = await client.me.messages.by_message_id(draft_id).attachments.create_upload_session.post(body)
        if not session or not session.upload_url:
            return ("session_error", None)
        return ("upload_url", session.upload_url)

    mode, value = run_graph(_run_graph_part())

    if mode == "not_draft":
        err_console.print(f"[red]Message {draft_id} is not a draft. Attachments can only be added to drafts.[/red]")
        raise typer.Exit(2)
    if mode == "session_error":
        err_console.print("[red]Upload session created but no upload URL was returned.[/red]")
        raise typer.Exit(1)

    attachment_id: str | None = None
    if mode == "inline":
        attachment_id = value
    else:  # upload_url
        _attach_chunked_put(value, file, file_size)

    payload = {
        "id": attachment_id,
        "name": display_name,
        "size": file_size,
        "contentType": content_type,
        "draftId": draft_id,
    }
    if as_json:
        import json as _json
        _json.dump(payload, sys.stdout)
        sys.stdout.write("\n")
        return

    err_console.print(
        f"[green]Attached[/green] {display_name} ({_format_bytes(file_size)}) to draft {draft_id}"
    )


# ── tmp ───────────────────────────────────────────────────────────────────

tmp_app = typer.Typer(help="Manage the ephemeral attachment download cache.", no_args_is_help=True)
app.add_typer(tmp_app, name="tmp")


@tmp_app.command("clean")
def tmp_clean(
    all_: Annotated[bool, typer.Option("--all", help="Wipe everything, ignoring the 24h TTL.")] = False,
):
    """Remove cached attachment downloads from ~/.outlook-cli/tmp."""
    if not _TMP_ROOT.exists():
        err_console.print("[dim]Nothing to clean.[/dim]")
        return

    import shutil

    removed = 0
    cutoff = 0 if all_ else _TMP_TTL_SECONDS
    now = _now()
    for child in _TMP_ROOT.iterdir():
        try:
            age = now - child.stat().st_mtime
        except OSError:
            continue
        if all_ or age > cutoff:
            shutil.rmtree(child, ignore_errors=True)
            removed += 1
    err_console.print(f"[green]Cleaned {removed}[/green] tmp entries from {_TMP_ROOT}")


def _tmp_dir_for_message(message_id: str) -> Path:
    """Return a private tmp dir for a message id, creating it if needed.

    Uses a short hash of the message id so the path stays sane on disk
    even though Graph message ids are long base64-ish strings.
    """
    import hashlib

    # Path.mkdir(parents=True, mode=...) does not apply mode to intermediates,
    # so create each level explicitly with 0700.
    _TMP_ROOT.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    _TMP_ROOT.mkdir(exist_ok=True, mode=0o700)
    digest = hashlib.sha256(message_id.encode()).hexdigest()[:16]
    target = _TMP_ROOT / digest
    target.mkdir(exist_ok=True, mode=0o700)
    return target


def _gc_tmp_root() -> None:
    """Opportunistically wipe tmp entries older than _TMP_TTL_SECONDS."""
    if not _TMP_ROOT.exists():
        return
    import shutil

    now = _now()
    for child in _TMP_ROOT.iterdir():
        try:
            age = now - child.stat().st_mtime
        except OSError:
            continue
        if age > _TMP_TTL_SECONDS:
            shutil.rmtree(child, ignore_errors=True)


def _now() -> float:
    import time
    return time.time()


def _attach_chunked_put(upload_url: str, file: Path, total_size: int) -> None:
    """PUT a file to a pre-authenticated upload URL in chunks."""
    _validate_graph_upload_url(upload_url)

    with file.open("rb") as fh:
        offset = 0
        while offset < total_size:
            chunk = fh.read(_UPLOAD_CHUNK_SIZE)
            if not chunk:
                break
            end = offset + len(chunk) - 1
            req = urllib.request.Request(
                upload_url,
                data=chunk,
                method="PUT",
                headers={
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {offset}-{end}/{total_size}",
                },
            )
            try:
                with urllib.request.urlopen(req) as resp:  # noqa: S310 — URL was allowlist-validated
                    if resp.status not in (200, 201, 202):
                        err_console.print(f"[red]Upload chunk failed: HTTP {resp.status}[/red]")
                        raise typer.Exit(1)
            except urllib.error.HTTPError as e:
                err_console.print(f"[red]Upload chunk failed: HTTP {e.code} {e.reason}[/red]")
                raise typer.Exit(1) from None
            offset += len(chunk)
            pct = int(offset / total_size * 100)
            err_console.print(f"\rUploading... {pct}%", end="")
        err_console.print()


def _download_one(att: Any, out_dir: Path, *, _skip_validate: bool = False) -> str:
    """Persist a single FileAttachment to disk. Returns the saved path as a string."""
    kind = _ODATA_KIND.get(att.odata_type or "", "unknown")
    if kind != "file":
        err_console.print(f"[red]Cannot download {kind} attachment {att.name or att.id}.[/red]")
        raise typer.Exit(2)
    raw = getattr(att, "content_bytes", None)
    if raw is None:
        err_console.print(f"[red]Attachment {att.name or att.id} has no content bytes.[/red]")
        raise typer.Exit(1)
    # Graph returns FileAttachment.contentBytes as a base64 string. The Python SDK
    # surfaces it as bytes but does NOT decode — we have to do that ourselves.
    try:
        content = base64.b64decode(raw, validate=True)
    except (binascii.Error, ValueError):
        # Already-decoded payload (defensive — observed on some SDK versions).
        content = raw
    if len(content) > _MAX_DOWNLOAD_BYTES:
        err_console.print(
            f"[red]Attachment {att.name or att.id} is {len(content)} bytes, exceeds "
            f"{_MAX_DOWNLOAD_BYTES} byte download cap.[/red]"
        )
        raise typer.Exit(1)
    if not _skip_validate:
        _validate_out_dir(out_dir)
    safe_name = _sanitize_filename(att.name or "attachment")
    return _safe_write_file(out_dir / safe_name, content)


def _validate_out_dir(path: Path) -> None:
    """Ensure the output directory exists, isn't a symlink, and is private (mode 0700)."""
    path.mkdir(parents=True, exist_ok=True, mode=0o700)
    info = path.lstat()
    if os.path.islink(path):
        err_console.print(f"[red]Output directory {path} is a symlink, refusing to write.[/red]")
        raise typer.Exit(2)
    from stat import S_ISDIR
    if not S_ISDIR(info.st_mode):
        err_console.print(f"[red]Output path {path} is not a directory.[/red]")
        raise typer.Exit(2)


def _safe_write_file(path: Path, content: bytes) -> str:
    """Write content to path, refusing to overwrite. On collision append (1), (2), ... up to 1000."""
    candidates = [path] + [path.with_name(f"{path.stem}({i}){path.suffix}") for i in range(1, 1000)]
    for candidate in candidates:
        try:
            fd = os.open(str(candidate), os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        except FileExistsError:
            continue
        with os.fdopen(fd, "wb") as fh:
            fh.write(content)
        return str(candidate)
    err_console.print(f"[red]Could not find an available filename near {path} after 1000 attempts.[/red]")
    raise typer.Exit(1)


def _sanitize_filename(name: str) -> str:
    """Strip path separators, control chars, and leading dots from an attachment-supplied name."""
    name = os.path.basename(name)
    name = name.replace(os.sep, "_").replace("/", "_").replace("\\", "_")
    name = "".join(ch for ch in name if ord(ch) >= 0x20 and ord(ch) != 0x7F)
    name = name.lstrip(".")
    return name or "attachment"


def _validate_graph_upload_url(raw_url: str) -> None:
    """Refuse pre-authenticated URLs that aren't HTTPS or aren't on a known Microsoft host."""
    parsed = urlparse(raw_url)
    if parsed.scheme != "https":
        err_console.print(f"[red]Refusing non-HTTPS upload URL ({parsed.scheme}).[/red]")
        raise typer.Exit(1)
    host = (parsed.hostname or "").lower()
    if not any(host.endswith(suffix) for suffix in _TRUSTED_UPLOAD_HOST_SUFFIXES):
        err_console.print(f"[red]Refusing upload URL on untrusted host {host!r}.[/red]")
        raise typer.Exit(1)


def _format_bytes(n: int | None) -> str:
    if n is None:
        return ""
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f}{unit}" if unit == "B" else f"{n:.1f}{unit}"
        n /= 1024
    return f"{n:.1f}GB"


# ── helpers ───────────────────────────────────────────────────────────────

def _resolve_body(body: str | None, body_file: Path | None, raw: bool = False) -> str:
    """Pick the right body source and (unless --raw-body) interpret escape sequences.

    Stdin and file sources are NOT escape-decoded — they're already real text.
    Only the `--body` arg gets decoded, because that's the path agents
    accidentally double-encode through.
    """
    if body is not None and body_file is not None:
        err_console.print("[red]--body and --body-file are mutually exclusive.[/red]")
        raise typer.Exit(2)
    if body == "-":
        return sys.stdin.read()
    if body is not None:
        return body if raw else interpret_escapes(body)
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


def _normalise_date(s: str) -> str:
    """Turn a YYYY-MM-DD or ISO 8601 string into the form Graph's $filter accepts.

    Graph wants ISO-8601 with a Z suffix inside $filter expressions.
    """
    s = s.strip()
    if "T" not in s:
        s += "T00:00:00"
    if not s.endswith("Z") and not s.endswith("z"):
        s += "Z"
    return s


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
