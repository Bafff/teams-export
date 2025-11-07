"""Interactive chat selection utilities."""

from __future__ import annotations

import sys
from typing import List, Sequence

import typer


def _chat_display_name(chat: dict) -> str:
    """Get a readable display name for a chat."""
    topic = chat.get("topic") or chat.get("displayName")
    if topic:
        return topic

    members = chat.get("members", [])
    if members:
        names = []
        for m in members:
            name = m.get("displayName") or m.get("email")
            if name:
                names.append(name)
        if names:
            return ", ".join(names)

    return chat.get("id", "Unknown chat")


def _chat_type_label(chat: dict) -> str:
    """Get a human-readable chat type label."""
    chat_type = chat.get("chatType", "").lower()
    if chat_type == "oneonone":
        return "1:1"
    elif chat_type == "group":
        return "Group"
    elif chat_type == "meeting":
        return "Meeting"
    return chat_type.title() if chat_type else "Unknown"


def _chat_last_updated(chat: dict) -> str:
    """Extract last updated timestamp for sorting."""
    return chat.get("lastUpdatedDateTime", "")


def select_chat_interactive(
    chats: Sequence[dict],
    *,
    prompt_message: str = "Select a chat:",
    show_limit: int = 20,
) -> dict:
    """Present an interactive menu to choose from multiple chats.

    Args:
        chats: List of chat objects to choose from
        prompt_message: Message to display before the menu
        show_limit: Maximum number of chats to show initially

    Returns:
        Selected chat object

    Raises:
        typer.Abort: If user cancels selection
    """

    if not chats:
        typer.secho("No chats available to select.", fg=typer.colors.RED)
        raise typer.Abort()

    if len(chats) == 1:
        return chats[0]

    # Sort by last updated (most recent first)
    sorted_chats = sorted(chats, key=_chat_last_updated, reverse=True)

    # Show up to show_limit chats
    display_chats = sorted_chats[:show_limit]

    typer.echo(f"\n{prompt_message}")
    typer.echo("=" * 80)
    typer.echo(f"{'#':<4} {'Type':<8} {'Chat Name':<50} {'Last Updated':<20}")
    typer.echo("-" * 80)

    for idx, chat in enumerate(display_chats, 1):
        name = _chat_display_name(chat)
        chat_type = _chat_type_label(chat)
        last_updated = chat.get("lastUpdatedDateTime", "N/A")

        # Truncate long names
        if len(name) > 47:
            name = name[:44] + "..."

        # Format timestamp
        if last_updated and last_updated != "N/A":
            try:
                # Show just date and time without milliseconds
                timestamp_display = last_updated.split('.')[0].replace('T', ' ')
            except Exception:
                timestamp_display = last_updated[:19]
        else:
            timestamp_display = "N/A"

        typer.echo(f"{idx:<4} {chat_type:<8} {name:<50} {timestamp_display:<20}")

    if len(sorted_chats) > show_limit:
        typer.echo("-" * 80)
        typer.echo(f"... and {len(sorted_chats) - show_limit} more chats (showing most recent {show_limit})")

    typer.echo("=" * 80)

    # Get user selection
    while True:
        try:
            selection = typer.prompt(
                f"\nEnter chat number (1-{len(display_chats)}) or 'q' to quit",
                default="",
            )

            if selection.lower() in ("q", "quit", "exit"):
                typer.echo("Selection cancelled.")
                raise typer.Abort()

            if not selection:
                continue

            choice = int(selection)
            if 1 <= choice <= len(display_chats):
                selected_chat = display_chats[choice - 1]
                selected_name = _chat_display_name(selected_chat)
                typer.secho(f"\n✓ Selected: {selected_name}", fg=typer.colors.GREEN)
                return selected_chat
            else:
                typer.secho(
                    f"Please enter a number between 1 and {len(display_chats)}.",
                    fg=typer.colors.YELLOW,
                )
        except ValueError:
            typer.secho("Invalid input. Please enter a number.", fg=typer.colors.YELLOW)
        except (KeyboardInterrupt, EOFError):
            typer.echo("\nSelection cancelled.")
            raise typer.Abort()


def filter_chats_by_query(chats: Sequence[dict], query: str) -> List[dict]:
    """Filter chats by a search query (case-insensitive substring match).

    Searches in:
    - Chat topic/display name
    - Member names
    - Member emails
    """
    if not query:
        return list(chats)

    query_lower = query.lower()
    matches = []

    for chat in chats:
        # Check chat name
        name = _chat_display_name(chat).lower()
        if query_lower in name:
            matches.append(chat)
            continue

        # Check members
        members = chat.get("members", [])
        for member in members:
            display_name = (member.get("displayName") or "").lower()
            email = (member.get("email") or "").lower()
            if query_lower in display_name or query_lower in email:
                matches.append(chat)
                break

    return matches
