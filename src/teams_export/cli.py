from __future__ import annotations

from pathlib import Path
from typing import Iterable

import typer

from .auth import AuthError, acquire_token
from .config import ConfigError, load_config
from .dates import DateParseError, resolve_range
from .exporter import ChatNotFoundError, choose_chat, export_chat
from .graph import GraphClient

app = typer.Typer(add_completion=False, help="Export Microsoft Teams chats via Microsoft Graph.")


def _chat_title(chat: dict) -> str:
    title = chat.get("topic") or chat.get("displayName")
    if title:
        return title
    members = chat.get("members", [])
    if members:
        return ", ".join(m.get("displayName") or m.get("email") or "?" for m in members)
    return chat.get("id", "<unknown chat>")


def _participants(chat: dict) -> str:
    members = chat.get("members", [])
    labels = []
    for member in members:
        label = member.get("displayName") or member.get("email")
        if label:
            labels.append(label)
    return ", ".join(labels)


def _print_chat_list(chats: Iterable[dict]) -> None:
    for chat in chats:
        typer.echo(
            f"{chat.get('id')}\t{chat.get('chatType')}\t{_chat_title(chat)}\t{_participants(chat)}"
        )


@app.command()
def main(
    participant: str = typer.Option(
        None,
        "--user",
        "-u",
        help="Participant display name or email for 1:1 chats.",
    ),
    chat_name: str = typer.Option(
        None,
        "--chat",
        "-c",
        help="Chat display name for group chats.",
    ),
    from_date: str = typer.Option(
        None,
        "--from",
        "-f",
        help='Start date (YYYY-MM-DD, "today", or "last week").',
    ),
    to_date: str = typer.Option(
        None,
        "--to",
        "-t",
        help='End date (YYYY-MM-DD, "today", or "last week").',
    ),
    output_format: str = typer.Option(
        "json",
        "--format",
        "-o",
        case_sensitive=False,
        help="Export format: json or csv.",
    ),
    output_dir: Path = typer.Option(
        Path("exports"),
        "--output-dir",
        show_default=True,
        help="Directory to save exported files.",
    ),
    list_chats: bool = typer.Option(
        False,
        "--list",
        help="List accessible chats and exit.",
    ),
    export_all: bool = typer.Option(
        False,
        "--all",
        help="Export all chats within the date range.",
    ),
    force_login: bool = typer.Option(
        False,
        "--force-login",
        help="Skip cache and refresh the device login flow.",
    ),
) -> None:
    try:
        config = load_config()
    except ConfigError as exc:
        typer.secho(f"Configuration error: {exc}", fg=typer.colors.RED)
        raise typer.Exit(code=1)

    try:
        start_dt, end_dt = resolve_range(from_date, to_date)
    except DateParseError as exc:
        typer.secho(f"Invalid date input: {exc}", fg=typer.colors.RED)
        raise typer.Exit(code=2)

    typer.echo("Authenticating with Microsoft Graph…")
    try:
        token = acquire_token(config, message_callback=typer.echo, force_refresh=force_login)
    except AuthError as exc:
        typer.secho(f"Authentication failed: {exc}", fg=typer.colors.RED)
        raise typer.Exit(code=3)

    with GraphClient(token) as client:
        chats = client.list_chats()
        if list_chats:
            typer.echo("Chat ID\tType\tTitle\tParticipants")
            _print_chat_list(chats)
            raise typer.Exit()

        exports: list[tuple[str, Path, int]] = []

        if export_all:
            selected_chats = chats
        else:
            if not participant and not chat_name:
                prompt_value = typer.prompt("Enter chat partner name/email (leave blank to use chat name)", default="")
                if prompt_value:
                    participant = prompt_value
                else:
                    chat_name = typer.prompt("Enter chat display name", default="") or None
            try:
                chat = choose_chat(chats, participant=participant, chat_name=chat_name)
            except ChatNotFoundError as exc:
                typer.secho(str(exc), fg=typer.colors.RED)
                raise typer.Exit(code=4)
            selected_chats = [chat]

        total_messages = 0
        for chat in selected_chats:
            title = _chat_title(chat)
            typer.echo(f"Exporting chat: {title}")
            try:
                output_path, count = export_chat(
                    client,
                    chat,
                    start_dt,
                    end_dt,
                    output_dir=output_dir,
                    output_format=output_format,
                )
            except ValueError as exc:
                typer.secho(str(exc), fg=typer.colors.RED)
                raise typer.Exit(code=5)

            exports.append((title, output_path, count))
            total_messages += count

    for title, path, count in exports:
        typer.echo(f"Exported {count} messages from {title}; saved to {path}")

    typer.echo(
        f"✅ Export complete. Total messages: {total_messages}. Date range: {start_dt.date()} to {end_dt.date()}"
    )
