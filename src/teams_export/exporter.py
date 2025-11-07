from __future__ import annotations

import csv
import json
import re
from pathlib import Path
from typing import Iterable, List, Sequence
from urllib.parse import urlparse

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
    from_field = message.get("from") or {}
    sender_info = from_field.get("user") or {}
    sender_fallback = from_field.get("application") or {}
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


def _download_attachment(client: GraphClient, url: str, output_path: Path) -> bool:
    """Download an attachment from a URL to local file.

    Returns:
        True if download succeeded, False otherwise
    """
    try:
        # Use the authenticated session from GraphClient
        response = client._session.get(url, timeout=30)
        if response.status_code == 200:
            output_path.write_bytes(response.content)
            return True
        else:
            print(f"Failed to download {url}: HTTP {response.status_code}")
            return False
    except Exception as e:
        print(f"Error downloading {url}: {e}")
        return False


def _extract_image_urls(messages: Sequence[dict]) -> List[str]:
    """Extract all image URLs from messages (both inline and attachments)."""
    import re

    urls = []
    for message in messages:
        # Extract inline images from HTML content
        content = message.get("content", "")
        if content:
            img_pattern = r'<img[^>]+src=["\']([^"\']+)["\'][^>]*>'
            for match in re.finditer(img_pattern, content, flags=re.IGNORECASE):
                url = match.group(1)
                if url and url.startswith("http"):
                    urls.append(url)

        # Extract from attachments array
        attachments = message.get("attachments", [])
        for att in attachments:
            # Try different possible URL fields
            url = (
                att.get("contentUrl") or
                att.get("content") or
                att.get("url") or
                att.get("thumbnailUrl") or
                (att.get("hostedContents", {}).get("contentUrl") if isinstance(att.get("hostedContents"), dict) else None)
            )
            if url and url.startswith("http"):
                # Check if it's an image
                content_type = att.get("contentType", "")
                name = att.get("name", "")
                is_image = (
                    content_type.startswith("image/") if content_type else
                    any(name.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp'])
                )
                if is_image:
                    urls.append(url)

    return urls


def _download_attachments(
    client: GraphClient,
    messages: Sequence[dict],
    attachments_dir: Path,
) -> dict[str, str]:
    """Download all image attachments and return URL -> local path mapping.

    Args:
        client: Authenticated Graph API client
        messages: List of message dictionaries
        attachments_dir: Directory to save attachments

    Returns:
        Dictionary mapping original URL to local relative path
    """
    attachments_dir.mkdir(parents=True, exist_ok=True)

    urls = _extract_image_urls(messages)
    unique_urls = list(dict.fromkeys(urls))  # Remove duplicates while preserving order

    url_mapping = {}

    if not unique_urls:
        return url_mapping

    print(f"\nDownloading {len(unique_urls)} image(s)...")

    for idx, url in enumerate(unique_urls, 1):
        # Generate filename from URL or use index
        try:
            parsed = urlparse(url)
            path_parts = parsed.path.split('/')
            # Try to get a meaningful name from the URL
            if path_parts:
                filename = path_parts[-1] or f"image_{idx:03d}"
            else:
                filename = f"image_{idx:03d}"

            # If no extension, try to guess from URL or default to .png
            if '.' not in filename:
                if 'image' in url.lower():
                    filename += '.png'
                else:
                    filename += '.bin'
        except Exception:
            filename = f"image_{idx:03d}.png"

        # Sanitize filename
        filename = re.sub(r'[^\w\-.]', '_', filename)
        output_path = attachments_dir / filename

        # Avoid overwriting if file already exists
        counter = 1
        while output_path.exists():
            name_part = filename.rsplit('.', 1)[0]
            ext_part = '.' + filename.rsplit('.', 1)[1] if '.' in filename else ''
            output_path = attachments_dir / f"{name_part}_{counter}{ext_part}"
            counter += 1

        if _download_attachment(client, url, output_path):
            # Store relative path (relative to the markdown file)
            relative_path = f"{attachments_dir.name}/{output_path.name}"
            url_mapping[url] = relative_path
            print(f"  [{idx}/{len(unique_urls)}] Downloaded: {output_path.name}")
        else:
            print(f"  [{idx}/{len(unique_urls)}] Failed: {url}")

    return url_mapping


def export_chat(
    client: GraphClient,
    chat: dict,
    start_dt,
    end_dt,
    *,
    output_dir: Path,
    output_format: str = "json",
    download_attachments: bool = True,
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

    # Download attachments if requested (only for formats that support it)
    url_mapping = {}
    if download_attachments and fmt == "jira" and messages:
        # Create attachments directory next to output file
        attachments_dir_name = output_path.stem + "_files"
        attachments_dir = output_path.parent / attachments_dir_name
        url_mapping = _download_attachments(client, messages, attachments_dir)

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
        write_jira_markdown(messages, output_path, chat_info=chat_info, url_mapping=url_mapping)
    else:
        raise ValueError("Unsupported export format. Choose json, csv, or jira.")

    return output_path, message_count
