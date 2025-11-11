from __future__ import annotations

import csv
import json
import re
from pathlib import Path
from typing import Iterable, List, Sequence
from urllib.parse import urlparse

from dateutil import parser

from .graph import GraphClient
from .formatters import write_jira_markdown, write_html, write_docx


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


def _get_extension_from_mime(mime_type: str) -> str:
    """Get file extension from MIME type."""
    mime_to_ext = {
        # Images
        'image/png': '.png',
        'image/jpeg': '.jpg',
        'image/jpg': '.jpg',
        'image/gif': '.gif',
        'image/bmp': '.bmp',
        'image/webp': '.webp',
        'image/svg+xml': '.svg',
        'image/tiff': '.tiff',
        'image/x-icon': '.ico',
        # Documents
        'application/pdf': '.pdf',
        'application/msword': '.doc',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
        'application/vnd.ms-excel': '.xls',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
        'application/vnd.ms-powerpoint': '.ppt',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
        'application/vnd.oasis.opendocument.text': '.odt',
        'application/vnd.oasis.opendocument.spreadsheet': '.ods',
        'application/vnd.oasis.opendocument.presentation': '.odp',
        # Archives
        'application/zip': '.zip',
        'application/x-zip-compressed': '.zip',
        'application/x-rar-compressed': '.rar',
        'application/x-7z-compressed': '.7z',
        'application/gzip': '.gz',
        'application/x-tar': '.tar',
        # Text
        'text/plain': '.txt',
        'text/csv': '.csv',
        'text/html': '.html',
        'text/css': '.css',
        'text/javascript': '.js',
        'application/json': '.json',
        'application/xml': '.xml',
        'text/xml': '.xml',
        'text/markdown': '.md',
        # Code
        'application/x-python': '.py',
        'text/x-python': '.py',
        'application/x-sh': '.sh',
        # Video
        'video/mp4': '.mp4',
        'video/mpeg': '.mpeg',
        'video/quicktime': '.mov',
        'video/x-msvideo': '.avi',
        'video/webm': '.webm',
        # Audio
        'audio/mpeg': '.mp3',
        'audio/wav': '.wav',
        'audio/ogg': '.ogg',
        'audio/webm': '.weba',
    }
    return mime_to_ext.get(mime_type.lower(), '.bin')


def _download_attachment(client: GraphClient, url: str, output_path: Path) -> tuple[bool, str | None]:
    """Download an attachment from a URL to local file.

    Returns:
        Tuple of (success: bool, content_type: str | None)
    """
    try:
        # Use the authenticated session from GraphClient
        response = client._session.get(url, timeout=30)
        if response.status_code == 200:
            output_path.write_bytes(response.content)
            content_type = response.headers.get('Content-Type', '').split(';')[0].strip()
            return True, content_type
        else:
            print(f"Failed to download {url}: HTTP {response.status_code}")
            return False, None
    except Exception as e:
        print(f"Error downloading {url}: {e}")
        return False, None


def _extract_attachment_urls(messages: Sequence[dict]) -> List[tuple[str, bool]]:
    """Extract all attachment URLs from messages (both inline images and file attachments).

    Returns:
        List of tuples (url, is_image) where is_image indicates if the attachment is an image.
    """
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
                    urls.append((url, True))  # Inline images are always images

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
                urls.append((url, is_image))

    return urls


def _download_attachments(
    client: GraphClient,
    messages: Sequence[dict],
    attachments_dir: Path,
) -> dict[str, str]:
    """Download all attachments (images and files) and return URL -> local path mapping.

    Args:
        client: Authenticated Graph API client
        messages: List of message dictionaries
        attachments_dir: Directory to save attachments

    Returns:
        Dictionary mapping original URL to local relative path
    """
    attachments_dir.mkdir(parents=True, exist_ok=True)

    url_tuples = _extract_attachment_urls(messages)
    unique_url_tuples = list(dict.fromkeys(url_tuples))  # Remove duplicates while preserving order

    url_mapping = {}

    if not unique_url_tuples:
        return url_mapping

    # Count images and non-images
    image_count = sum(1 for _, is_image in unique_url_tuples if is_image)
    file_count = len(unique_url_tuples) - image_count

    if image_count and file_count:
        print(f"\nDownloading {image_count} image(s) and {file_count} file(s)...")
    elif image_count:
        print(f"\nDownloading {image_count} image(s)...")
    else:
        print(f"\nDownloading {file_count} file(s)...")

    for idx, (url, is_image) in enumerate(unique_url_tuples, 1):
        # Generate base filename (without extension) from URL or use index
        try:
            parsed = urlparse(url)
            path_parts = parsed.path.split('/')
            # Try to get a meaningful name from the URL
            if path_parts and path_parts[-1]:
                base_filename = path_parts[-1]
                # Remove extension if present, we'll add correct one later
                if '.' in base_filename:
                    base_filename = base_filename.rsplit('.', 1)[0]
            else:
                # Use appropriate prefix based on file type
                prefix = "image" if is_image else "file"
                base_filename = f"{prefix}_{idx:03d}"
        except Exception:
            # Use appropriate prefix based on file type
            prefix = "image" if is_image else "file"
            base_filename = f"{prefix}_{idx:03d}"

        # Sanitize base filename
        base_filename = re.sub(r'[^\w\-]', '_', base_filename)

        # Download to temporary path first to get Content-Type
        temp_filename = f"{base_filename}_temp"
        temp_path = attachments_dir / temp_filename

        success, content_type = _download_attachment(client, url, temp_path)

        if success:
            # Determine correct extension from Content-Type
            if content_type:
                extension = _get_extension_from_mime(content_type)
            else:
                # Fallback based on type
                extension = '.png' if is_image else '.bin'

            # Create final filename with correct extension
            final_filename = f"{base_filename}{extension}"
            final_path = attachments_dir / final_filename

            # Avoid overwriting if file already exists
            counter = 1
            while final_path.exists():
                final_filename = f"{base_filename}_{counter}{extension}"
                final_path = attachments_dir / final_filename
                counter += 1

            # Rename from temp to final name
            temp_path.rename(final_path)

            # Store relative path (relative to the markdown file)
            relative_path = f"{attachments_dir.name}/{final_path.name}"
            url_mapping[url] = relative_path
            print(f"  [{idx}/{len(unique_urls)}] Downloaded: {final_path.name}")
        else:
            # Clean up temp file if exists
            if temp_path.exists():
                temp_path.unlink()
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
    elif fmt == "html":
        suffix = "html"
    elif fmt in ("docx", "word"):
        suffix = "docx"
        fmt = "docx"
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
    attachments_dir = None
    if download_attachments and fmt in ("jira", "html", "docx") and messages:
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
    elif fmt == "html":
        # Prepare chat metadata for HTML formatter
        chat_title = chat.get("topic") or chat.get("displayName") or identifier
        participants_list = _member_labels(chat)
        chat_info = {
            "title": chat_title,
            "participants": ", ".join(participants_list) if participants_list else "N/A",
            "date_range": f"{start_dt.date()} to {end_dt.date()}",
        }
        write_html(messages, output_path, chat_info=chat_info, url_mapping=url_mapping)
    elif fmt == "docx":
        # Prepare chat metadata for Word document formatter
        chat_title = chat.get("topic") or chat.get("displayName") or identifier
        participants_list = _member_labels(chat)
        chat_info = {
            "title": chat_title,
            "participants": ", ".join(participants_list) if participants_list else "N/A",
            "date_range": f"{start_dt.date()} to {end_dt.date()}",
        }
        write_docx(messages, output_path, chat_info=chat_info, url_mapping=url_mapping)
    else:
        raise ValueError("Unsupported export format. Choose json, csv, jira, html, or docx.")

    return output_path, message_count
