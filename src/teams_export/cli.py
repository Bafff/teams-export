from __future__ import annotations

import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable

import typer

from .auth import AuthError, acquire_token
from .cache import ChatCache
from .config import ConfigError, load_config
from .dates import DateParseError, resolve_range
from .exporter import ChatNotFoundError, choose_chat, export_chat
from .graph import GraphClient
from .interactive import select_chat_interactive

app = typer.Typer(
    add_completion=False,
    help=(
        "Export Microsoft Teams chats via Microsoft Graph. "
        "Docs: https://arkadium.atlassian.net/wiki/spaces/IT/overview"
    ),
)


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


def _select_date_range_interactive() -> tuple[datetime, datetime] | None:
    """Interactively select date range for export.

    Returns:
        Tuple of (start_dt, end_dt), or None if user cancels
    """
    typer.echo("\n📅 Select export period:")
    typer.echo("  1. Today (last 24 hours)")
    typer.echo("  2. Last 7 days")
    typer.echo("  3. Last 30 days")
    typer.echo("  4. Last 90 days")
    typer.echo("  5. All time (last 1 year)")
    typer.echo("  6. Custom date range")
    typer.echo("  q. Cancel export")

    choice = typer.prompt("\nYour choice", default="1").strip().lower()

    if choice == "q":
        return None

    now = datetime.now().astimezone()

    if choice == "1":
        # Last 24 hours (today)
        start_dt = now - timedelta(hours=24)
        end_dt = now
    elif choice == "2":
        start_dt = now - timedelta(days=7)
        end_dt = now
    elif choice == "3":
        start_dt = now - timedelta(days=30)
        end_dt = now
    elif choice == "4":
        start_dt = now - timedelta(days=90)
        end_dt = now
    elif choice == "5":
        start_dt = now - timedelta(days=365)
        end_dt = now
    elif choice == "6":
        # Custom range
        from_str = typer.prompt("Start date (YYYY-MM-DD, 'today', or 'last week')")
        to_str = typer.prompt("End date (YYYY-MM-DD, 'today', or 'last week')", default="today")
        try:
            start_dt, end_dt = resolve_range(from_str, to_str)
        except DateParseError as exc:
            typer.secho(f"Invalid date: {exc}", fg=typer.colors.RED)
            return None
    else:
        typer.secho("Invalid choice", fg=typer.colors.RED)
        return None

    typer.echo(f"  ✓ Period: {start_dt.date()} to {end_dt.date()}")
    return start_dt, end_dt


def _load_chats_with_progress(client: GraphClient) -> list[dict]:
    """Load all chats from Microsoft Graph with progress indicator.

    Returns:
        List of chat objects
    """
    def show_progress(count: int) -> None:
        sys.stdout.write(f"\rLoading chats... {count} loaded")
        sys.stdout.flush()

    typer.echo("Loading chats from Microsoft Graph...")
    chats = client.list_chats(limit=None, progress_callback=show_progress)

    # Clear progress line
    if chats:
        sys.stdout.write("\r" + " " * 50 + "\r")
        sys.stdout.flush()
        typer.secho(f"✓ Loaded {len(chats)} chats", fg=typer.colors.GREEN)

    return chats


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
        "jira",
        "--format",
        "-o",
        case_sensitive=False,
        help="Export format: jira (markdown), html (copy-paste to Jira/Confluence), json, or csv.",
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
    refresh_cache: bool = typer.Option(
        False,
        "--refresh-cache",
        help="Force refresh of chat list cache.",
    ),
    download_attachments: bool = typer.Option(
        True,
        "--download-attachments/--no-download-attachments",
        help="Download images and attachments locally (default: enabled).",
    ),
) -> None:
    try:
        config = load_config()
    except ConfigError as exc:
        typer.secho(f"Configuration error: {exc}", fg=typer.colors.RED)
        raise typer.Exit(code=1)

    # Parse dates if provided, otherwise they'll be set interactively
    start_dt = None
    end_dt = None
    if from_date or to_date:
        try:
            start_dt, end_dt = resolve_range(from_date, to_date)
        except DateParseError as exc:
            typer.secho(f"Invalid date input: {exc}", fg=typer.colors.RED)
            raise typer.Exit(code=2)

    typer.echo("Authenticating with Microsoft Graph…")
    try:
        token = acquire_token(config, message_callback=typer.echo, force_refresh=force_login)
        typer.secho("✓ Authenticated successfully", fg=typer.colors.GREEN)
    except AuthError as exc:
        typer.secho(f"Authentication failed: {exc}", fg=typer.colors.RED)
        raise typer.Exit(code=3)

    with GraphClient(token) as client:
        # Try to load from cache first
        cache = ChatCache()
        user_id = "me"  # Simple identifier for caching
        chats = None

        if not refresh_cache:
            chats = cache.get(user_id)
            if chats:
                typer.secho(f"✓ Loaded {len(chats)} chats from cache (24h TTL, press 'c' in menu to refresh)", fg=typer.colors.CYAN)

        # If no cache or refresh requested, load from API
        if chats is None:
            chats = _load_chats_with_progress(client)
            # Save to cache for next time
            if chats:
                cache.set(user_id, chats)

        if list_chats:
            typer.echo("\nChat ID\tType\tTitle\tParticipants")
            _print_chat_list(chats)
            raise typer.Exit()

        # Check if we found any chats
        if not chats:
            typer.secho("No chats found.", fg=typer.colors.YELLOW)
            raise typer.Exit(code=0)

        exports: list[tuple[str, Path, int]] = []

        # Set default date range if not provided (for --all and --user/--chat modes)
        if start_dt is None and (export_all or participant or chat_name):
            # Default to last 24 hours (today) if dates not specified
            now = datetime.now().astimezone()
            start_dt = now - timedelta(hours=24)
            end_dt = now
            typer.echo(f"Using default date range: last 24 hours ({start_dt.date()} to {end_dt.date()})")

        if export_all:
            selected_chats = chats
        else:
            if not participant and not chat_name:
                # Interactive mode - show chat menu (with cache refresh support)
                while True:
                    try:
                        chat = select_chat_interactive(
                            chats,
                            prompt_message="Select a chat to export:",
                            showing_limited=False,
                        )

                        # Check if user requested cache refresh
                        if isinstance(chat, dict) and chat.get("__action__") == "refresh_cache":
                            chats = _load_chats_with_progress(client)
                            if chats:
                                cache.set(user_id, chats)
                            continue  # Show menu again with refreshed data

                        # Ask for date range if not provided
                        if start_dt is None:
                            date_range = _select_date_range_interactive()
                            if date_range is None:
                                typer.echo("Export cancelled")
                                raise typer.Exit(code=0)
                            start_dt, end_dt = date_range

                        selected_chats = [chat]
                        break
                    except typer.Abort:
                        raise typer.Exit(code=0)
            else:
                # Search mode - try to find by participant or chat name
                try:
                    result = choose_chat(chats, participant=participant, chat_name=chat_name)
                except ChatNotFoundError as exc:
                    typer.secho(str(exc), fg=typer.colors.RED)
                    raise typer.Exit(code=4)

                # If multiple matches, let user choose interactively
                if isinstance(result, list):
                    typer.echo(f"\nFound {len(result)} matching chats.")
                    try:
                        chat = select_chat_interactive(
                            result,
                            prompt_message="Multiple chats matched. Please select one:",
                        )
                        selected_chats = [chat]
                    except typer.Abort:
                        raise typer.Exit(code=0)
                else:
                    selected_chats = [result]

        total_messages = 0

        # Use parallel processing for multiple chats
        if len(selected_chats) > 1:
            typer.echo(f"\nExporting {len(selected_chats)} chats in parallel...")

            def export_single_chat(chat):
                title = _chat_title(chat)
                try:
                    output_path, count = export_chat(
                        client,
                        chat,
                        start_dt,
                        end_dt,
                        output_dir=output_dir,
                        output_format=output_format,
                        download_attachments=download_attachments,
                    )
                    return (title, output_path, count, None)
                except Exception as exc:
                    return (title, None, 0, str(exc))

            # Use ThreadPoolExecutor for parallel downloads (limited to 3 concurrent)
            with ThreadPoolExecutor(max_workers=3) as executor:
                futures = {executor.submit(export_single_chat, chat): chat for chat in selected_chats}

                completed = 0
                for future in as_completed(futures):
                    title, output_path, count, error = future.result()
                    completed += 1

                    if error:
                        typer.secho(f"[{completed}/{len(selected_chats)}] Failed: {title} - {error}", fg=typer.colors.RED)
                    else:
                        exports.append((title, output_path, count))
                        total_messages += count
                        typer.secho(
                            f"[{completed}/{len(selected_chats)}] Exported {count} messages from {title}",
                            fg=typer.colors.GREEN
                        )
        else:
            # Single chat - process directly
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
                        download_attachments=download_attachments,
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
