from __future__ import annotations

import csv
import json
import re
from pathlib import Path
from typing import Iterable, List, Sequence

from dateutil import parser

from .graph import GraphClient
from .formatters import write_jira_markdown


class ChatNotFoundError(RuntimeError):
    """Raised when a chat matching the requested criteria cannot be found."""


def _normalise(value: str | None) -> str:
    return re.sub(r"\s+", " ", value or "").strip().lower()


def _member_labels(chat: dict) -> List[str]:
    labels: List[str] = []
    for member in chat.get("members", []):
        display = member.get("displayName")
        mail = member.get("email")
        if display:
            labels.append(display)
        if mail:
            labels.append(mail)
    return labels


def choose_chat(
    chats: Sequence[dict],
    *,
    participant: str | None = None,
    chat_name: str | None = None,
) -> dict | List[dict]:
    """Select a chat by participant identifier or chat display name.

    Returns:
        Either a single chat dict if exactly one match, or a list of matches
        if multiple chats matched the criteria.
    """

    name_norm = _normalise(chat_name) if chat_name else None
    participant_norm = _normalise(participant) if participant else None

    matches: List[dict] = []

    for chat in chats:
        chat_type = chat.get("chatType")
        topic = chat.get("topic") or chat.get("displayName")
        chat_label = _normalise(topic)
        if name_norm and chat_label == name_norm:
            matches.append(chat)
            continue
        if participant_norm:
            if chat_type and chat_type.lower() != "oneonone":
                continue
            for label in _member_labels(chat):
                if _normalise(label) == participant_norm:
                    matches.append(chat)
                    break

    if not matches:
        raise ChatNotFoundError(
            "No chat matches the provided identifiers. Try running with --list to"
            " review available chats."
        )
    if len(matches) == 1:
        return matches[0]

    # Return all matches for interactive selection
    return matches


def _normalise_filename(identifier: str) -> str:
    safe = re.sub(r"[^a-zA-Z0-9_-]+", "_", identifier.strip())
    return safe.lower().strip("_") or "chat"


def _transform_message(message: dict) -> dict:
    sender_info = message.get("from", {}).get("user", {})
    sender_fallback = message.get("from", {}).get("application", {})
    sender_display = sender_info.get("displayName") or sender_fallback.get("displayName")
    sender_email = sender_info.get("userPrincipalName") or sender_info.get("email")

    timestamp = message.get("lastModifiedDateTime") or message.get("createdDateTime")

    transformed = {
        "id": message.get("id"),
        "sender": sender_display,
        "sender_email": sender_email,
        "timestamp": timestamp,
        "type": message.get("messageType"),
        "subject": message.get("subject"),
        "content_type": message.get("body", {}).get("contentType"),
        "content": message.get("body", {}).get("content"),
        "reactions": message.get("reactions", []),
        "mentions": message.get("mentions", []),
        "attachments": message.get("attachments", []),
    }
    return transformed


def _within_range(message: dict, start_dt, end_dt) -> bool:
    timestamp = (
        message.get("lastModifiedDateTime")
        or message.get("createdDateTime")
        or message.get("originalArrivalDateTime")
    )
    if not timestamp:
        return False
    try:
        dt_value = parser.isoparse(timestamp)
    except (ValueError, TypeError):
        return False
    return start_dt <= dt_value <= end_dt


def _write_json(messages: Sequence[dict], output_path: Path) -> None:
    payload = list(messages)
    output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _write_csv(messages: Sequence[dict], output_path: Path) -> None:
    fieldnames = ["timestamp", "sender", "sender_email", "content", "type"]
    with output_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for message in messages:
            writer.writerow({key: message.get(key) for key in fieldnames})


def export_chat(
    client: GraphClient,
    chat: dict,
    start_dt,
    end_dt,
    *,
    output_dir: Path,
    output_format: str = "json",
) -> tuple[Path, int]:
    chat_id = chat.get("id")
    if not chat_id:
        raise ChatNotFoundError("Selected chat missing identifier.")

    identifier = chat.get("topic") or chat.get("displayName")
    if not identifier:
        members = _member_labels(chat)
        identifier = members[0] if members else chat_id
    filename_stem = _normalise_filename(identifier)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Normalize format and determine extension
    fmt = output_format.lower()
    if fmt in ("jira", "jira-markdown", "markdown"):
        suffix = "md"
        fmt = "jira"
    else:
        suffix = fmt

    if start_dt.date() == end_dt.date():
        date_fragment = start_dt.date().isoformat()
    else:
        date_fragment = f"{start_dt.date()}_{end_dt.date()}"
    output_path = output_dir / f"{filename_stem}_{date_fragment}.{suffix}"

    def _stop_condition(message: dict) -> bool:
        ts_value = message.get("createdDateTime") or message.get("lastModifiedDateTime")
        if not ts_value:
            return False
        try:
            dt_value = parser.isoparse(ts_value)
        except (ValueError, TypeError):
            return False
        return dt_value < start_dt

    raw_messages = client.list_chat_messages(chat_id, stop_condition=_stop_condition)
    filtered_messages = [m for m in raw_messages if _within_range(m, start_dt, end_dt)]

    # Sort messages from oldest to newest (Graph API returns newest first)
    filtered_messages.sort(
        key=lambda m: m.get("createdDateTime") or m.get("lastModifiedDateTime") or ""
    )

    messages = [_transform_message(m) for m in filtered_messages]
    message_count = len(messages)

    if fmt == "json":
        _write_json(messages, output_path)
    elif fmt == "csv":
        _write_csv(messages, output_path)
    elif fmt == "jira":
        # Prepare chat metadata for Jira formatter
        chat_title = chat.get("topic") or chat.get("displayName") or identifier
        participants_list = _member_labels(chat)
        chat_info = {
            "title": chat_title,
            "participants": ", ".join(participants_list) if participants_list else "N/A",
            "date_range": f"{start_dt.date()} to {end_dt.date()}",
        }
        write_jira_markdown(messages, output_path, chat_info=chat_info)
    else:
        raise ValueError("Unsupported export format. Choose json, csv, or jira.")

    return output_path, message_count
