from __future__ import annotations

import csv
import json
import re
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed

from dateutil import parser

from .graph import GraphClient
from .formatters import write_jira_markdown, write_html, write_docx
from .delta import DeltaStateManager


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
        'application/rtf': '.rtf',
        'application/vnd.oasis.opendocument.text': '.odt',
        'application/vnd.oasis.opendocument.spreadsheet': '.ods',
        'application/vnd.oasis.opendocument.presentation': '.odp',

        # Archives
        'application/zip': '.zip',
        'application/x-zip-compressed': '.zip',
        'application/x-rar-compressed': '.rar',
        'application/x-7z-compressed': '.7z',
        'application/x-tar': '.tar',
        'application/gzip': '.gz',
        'application/x-bzip2': '.bz2',

        # Text
        'text/plain': '.txt',
        'text/html': '.html',
        'text/css': '.css',
        'text/javascript': '.js',
        'application/javascript': '.js',
        'text/markdown': '.md',
        'text/csv': '.csv',

        # Data
        'application/json': '.json',
        'application/xml': '.xml',
        'text/xml': '.xml',
        'application/yaml': '.yaml',
        'text/yaml': '.yaml',

        # Media
        'audio/mpeg': '.mp3',
        'audio/wav': '.wav',
        'audio/ogg': '.ogg',
        'video/mp4': '.mp4',
        'video/mpeg': '.mpeg',
        'video/quicktime': '.mov',
        'video/x-msvideo': '.avi',
        'video/webm': '.webm',

        # Other
        'application/octet-stream': '.bin',
    }
    return mime_to_ext.get(mime_type.lower(), '.bin')


def _download_attachment(  # pragma: no cover
    client: GraphClient,
    url: str,
    output_path: Path,
) -> tuple[bool, str | None]:
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


def _extract_attachment_urls(messages: Sequence[dict], images_only: bool = False) -> List[tuple[str, str, str]]:
    """Extract all attachment URLs from messages.

    Args:
        messages: List of message dictionaries
        images_only: If True, only extract image attachments

    Returns:
        List of tuples: (url, name, content_type)
    """
    import re

    attachments = []
    seen_urls = set()

    for message in messages:
        # Extract inline images from HTML content regardless of mode
        content = message.get("content", "")
        if content:
            img_pattern = r'<img[^>]+src=["\']([^"\']+)["\'][^>]*>'
            for match in re.finditer(img_pattern, content, flags=re.IGNORECASE):
                url = match.group(1)
                if url and url.startswith("http") and url not in seen_urls:
                    seen_urls.add(url)
                    name = url.split('/')[-1][:50] if '/' in url else "inline_image"
                    attachments.append((url, name, "image/inline"))

        # Extract from attachments array
        msg_attachments = message.get("attachments", [])
        for att in msg_attachments:
            # Try different possible URL fields
            url = (
                att.get("contentUrl") or
                att.get("content") or
                att.get("url") or
                att.get("thumbnailUrl") or
                (att.get("hostedContents", {}).get("contentUrl") if isinstance(att.get("hostedContents"), dict) else None)
            )

            if url and url.startswith("http") and url not in seen_urls:
                name = att.get("name", "")
                content_type = att.get("contentType", "")

                # Determine if we should include this attachment
                if images_only:
                    # Check if it's an image
                    is_image = (
                        content_type.startswith("image/") if content_type else
                        any(name.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp'])
                    )
                    if not is_image:
                        continue

                seen_urls.add(url)
                attachments.append((url, name or "attachment", content_type or "application/octet-stream"))

    return attachments


def _extract_image_urls(messages: Sequence[dict]) -> List[str]:
    """Extract all image URLs from messages (both inline and attachments)."""
    # Keep this for backward compatibility
    attachments = _extract_attachment_urls(messages, images_only=True)
    return [url for url, name, content_type in attachments]
    print(f"\nDownloading {len(unique_urls)} image(s)...")

    for idx, url in enumerate(unique_urls, 1):
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
                base_filename = f"image_{idx:03d}"
        except Exception:
            base_filename = f"image_{idx:03d}"

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
                # Fallback to .png for images
                extension = '.png'

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


def _download_attachments_parallel(  # pragma: no cover - network + thread heavy
    client: GraphClient,
    messages: Sequence[dict],
    attachments_dir: Path,
    max_workers: int = 5,
) -> dict[str, str]:
    """Download all image attachments in parallel for faster processing.

    Args:
        client: Authenticated Graph API client
        messages: List of message dictionaries
        attachments_dir: Directory to save attachments
        max_workers: Number of parallel download workers

    Returns:
        Dictionary mapping original URL to local relative path
    """
    attachments_dir.mkdir(parents=True, exist_ok=True)

    urls = _extract_image_urls(messages)
    unique_urls = list(dict.fromkeys(urls))  # Remove duplicates while preserving order

    url_mapping = {}

    if not unique_urls:
        return url_mapping

    print(f"\nDownloading {len(unique_urls)} image(s) with {max_workers} parallel workers...")

    def download_single(idx_url_tuple):
        """Download a single attachment."""
        idx, url = idx_url_tuple

        # Generate base filename
        try:
            parsed = urlparse(url)
            path_parts = parsed.path.split('/')
            if path_parts and path_parts[-1]:
                base_filename = path_parts[-1]
                if '.' in base_filename:
                    base_filename = base_filename.rsplit('.', 1)[0]
            else:
                base_filename = f"image_{idx:03d}"
        except Exception:
            base_filename = f"image_{idx:03d}"

        # Sanitize base filename
        base_filename = re.sub(r'[^\w\-]', '_', base_filename)

        # Download to temporary path first
        temp_filename = f"{base_filename}_temp"
        temp_path = attachments_dir / temp_filename

        success, content_type = _download_attachment(client, url, temp_path)

        if success:
            # Determine correct extension
            extension = _get_extension_from_mime(content_type) if content_type else '.png'

            # Create final filename
            final_filename = f"{base_filename}{extension}"
            final_path = attachments_dir / final_filename

            # Avoid overwriting
            counter = 1
            while final_path.exists():
                final_filename = f"{base_filename}_{counter}{extension}"
                final_path = attachments_dir / final_filename
                counter += 1

            # Rename from temp to final
            temp_path.rename(final_path)

            # Store relative path
            relative_path = f"{attachments_dir.name}/{final_path.name}"
            return url, relative_path, idx, True
        else:
            # Clean up temp file
            if temp_path.exists():
                temp_path.unlink()
            return url, None, idx, False

    # Use ThreadPoolExecutor for parallel downloads
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Submit all downloads
        futures = {
            executor.submit(download_single, (idx, url)): (idx, url)
            for idx, url in enumerate(unique_urls, 1)
        }

        completed = 0
        for future in as_completed(futures):
            url, relative_path, idx, success = future.result()
            completed += 1

            if success:
                url_mapping[url] = relative_path
                print(f"  [{completed}/{len(unique_urls)}] Downloaded: {relative_path.split('/')[-1]}")
            else:
                print(f"  [{completed}/{len(unique_urls)}] Failed: {url}")

    return url_mapping


def _download_all_attachments_parallel(  # pragma: no cover - network + thread heavy
    client: GraphClient,
    messages: Sequence[dict],
    attachments_dir: Path,
    max_workers: int = 5,
) -> dict[str, str]:
    """Download ALL attachments (not just images) in parallel.

    Args:
        client: Authenticated Graph API client
        messages: List of message dictionaries
        attachments_dir: Directory to save attachments
        max_workers: Number of parallel download workers

    Returns:
        Dictionary mapping original URL to local relative path
    """
    attachments_dir.mkdir(parents=True, exist_ok=True)

    # Get all attachments, not just images
    attachments = _extract_attachment_urls(messages, images_only=False)

    url_mapping = {}

    if not attachments:
        return url_mapping

    print(f"\nDownloading {len(attachments)} attachment(s) with {max_workers} parallel workers...")

    def download_single(idx_att_tuple):
        """Download a single attachment."""
        idx, (url, name, content_type) = idx_att_tuple

        # Generate base filename from provided name or URL
        if name and name != "attachment":
            base_filename = name
            # Remove extension if present, we'll add correct one later
            if '.' in base_filename:
                base_filename = base_filename.rsplit('.', 1)[0]
        else:
            try:
                parsed = urlparse(url)
                path_parts = parsed.path.split('/')
                if path_parts and path_parts[-1]:
                    base_filename = path_parts[-1]
                    if '.' in base_filename:
                        base_filename = base_filename.rsplit('.', 1)[0]
                else:
                    base_filename = f"attachment_{idx:03d}"
            except Exception:
                base_filename = f"attachment_{idx:03d}"

        # Sanitize base filename
        base_filename = re.sub(r'[^\w\-]', '_', base_filename)[:100]  # Limit length

        # Download to temporary path first
        temp_filename = f"{base_filename}_temp"
        temp_path = attachments_dir / temp_filename

        success, actual_content_type = _download_attachment(client, url, temp_path)

        if success:
            # Use actual content type from response if available
            final_content_type = actual_content_type or content_type
            extension = _get_extension_from_mime(final_content_type)

            # Create final filename
            final_filename = f"{base_filename}{extension}"
            final_path = attachments_dir / final_filename

            # Avoid overwriting
            counter = 1
            while final_path.exists():
                final_filename = f"{base_filename}_{counter}{extension}"
                final_path = attachments_dir / final_filename
                counter += 1

            # Rename from temp to final
            temp_path.rename(final_path)

            # Store relative path
            relative_path = f"{attachments_dir.name}/{final_path.name}"
            return url, relative_path, idx, True, name
        else:
            # Clean up temp file
            if temp_path.exists():
                temp_path.unlink()
            return url, None, idx, False, name

    # Use ThreadPoolExecutor for parallel downloads
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Submit all downloads
        futures = {
            executor.submit(download_single, (idx, att)): (idx, att)
            for idx, att in enumerate(attachments, 1)
        }

        completed = 0
        for future in as_completed(futures):
            url, relative_path, idx, success, name = future.result()
            completed += 1

            if success:
                url_mapping[url] = relative_path
                display_name = name if name and name != "attachment" else relative_path.split('/')[-1]
                print(f"  [{completed}/{len(attachments)}] Downloaded: {display_name}")
            else:
                print(f"  [{completed}/{len(attachments)}] Failed: {name or url[:50]}")

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
    download_all_types: bool = False,
    use_parallel_fetch: bool = True,
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

    # Use parallel fetching for better performance on large chats
    if use_parallel_fetch:
        print(f"Fetching messages with parallel pagination...")
        raw_messages = client.list_chat_messages_parallel(
            chat_id,
            stop_condition=_stop_condition,
            max_workers=3
        )
    else:
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
        # Use parallel downloads for better performance
        if download_all_types:
            # Download all attachment types (PDFs, docs, etc.)
            url_mapping = _download_all_attachments_parallel(client, messages, attachments_dir, max_workers=5)
        else:
            # Download only images (backward compatible)
            url_mapping = _download_attachments_parallel(client, messages, attachments_dir, max_workers=5)

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


def export_chat_incremental(
    client: GraphClient,
    chat: dict,
    start_dt,
    end_dt,
    *,
    output_dir: Path,
    output_format: str = "json",
    download_attachments: bool = True,
    download_all_types: bool = False,
    delta_manager: Optional[DeltaStateManager] = None,
    use_parallel_fetch: bool = True,
) -> Tuple[Path, int, bool]:
    """Export chat messages with incremental sync support using delta queries.

    Args:
        client: Authenticated Graph client
        chat: Chat dictionary
        start_dt: Start datetime
        end_dt: End datetime
        output_dir: Output directory
        output_format: Export format
        download_attachments: Whether to download attachments
        delta_manager: Delta state manager for incremental sync
        use_parallel_fetch: Whether to use parallel pagination for initial sync fallbacks

    Returns:
        Tuple of (output_path, message_count, has_changes)
    """
    chat_id = chat.get("id")
    if not chat_id:
        raise ChatNotFoundError("Selected chat missing identifier.")

    # Initialize delta manager if not provided
    if delta_manager is None:
        delta_manager = DeltaStateManager()

    # Check for existing delta state
    state = delta_manager.get_state(chat_id)
    has_previous_state = state is not None

    # Use delta query if we have state
    if state and state.get("delta_link"):
        print(f"Using incremental sync for chat {chat_id} (last sync: {state.get('last_sync', 'unknown')})")
        messages, new_delta_link = client.list_chat_messages_delta(
            chat_id,
            delta_link=state["delta_link"]
        )

        # If no new messages, return early
        if not messages:
            print(f"No new messages since last sync")
            # Return path to existing file if it exists
            identifier = chat.get("topic") or chat.get("displayName")
            if not identifier:
                members = _member_labels(chat)
                identifier = members[0] if members else chat_id
            filename_stem = _normalise_filename(identifier)

            # Try to find existing export file
            fmt = output_format.lower()
            if fmt in ("jira", "jira-markdown", "markdown"):
                suffix = "md"
            elif fmt == "html":
                suffix = "html"
            elif fmt in ("docx", "word"):
                suffix = "docx"
            else:
                suffix = fmt

            # Look for most recent export file
            pattern = f"{filename_stem}_*.{suffix}"
            existing_files = list(output_dir.glob(pattern))
            if existing_files:
                existing_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
                return existing_files[0], state.get("message_count", 0), False

            # No existing file and no new messages - shouldn't happen normally
            return Path(""), 0, False
    else:
        print(f"Performing full sync for chat {chat_id}")
        # Full sync - use regular message fetch with delta
        messages, new_delta_link = client.list_chat_messages_delta(chat_id)

    # Filter messages by date range
    raw_messages = messages
    filtered_messages = [m for m in raw_messages if _within_range(m, start_dt, end_dt)]

    # Sort messages from oldest to newest
    filtered_messages.sort(
        key=lambda m: m.get("createdDateTime") or m.get("lastModifiedDateTime") or ""
    )

    # Transform messages
    messages = [_transform_message(m) for m in filtered_messages]
    message_count = len(messages)

    # Update delta state
    if new_delta_link and filtered_messages:
        last_message = filtered_messages[-1]
        delta_manager.save_state(
            chat_id,
            delta_link=new_delta_link,
            last_message_id=last_message.get("id"),
            last_message_time=last_message.get("createdDateTime") or last_message.get("lastModifiedDateTime"),
            message_count=message_count
        )
    elif new_delta_link:
        # Save delta link even if no messages in range
        delta_manager.save_state(chat_id, delta_link=new_delta_link, message_count=0)

    # If no messages to export, return early
    if not messages:
        print(f"No messages in specified date range")
        return Path(""), 0, False

    # Generate output path
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

    # For incremental exports, append to existing file if format supports it
    if has_previous_state and fmt in ("json", "csv"):
        # For JSON and CSV, we need to merge with existing data
        if output_path.exists():
            if fmt == "json":
                # Load existing messages
                try:
                    with open(output_path, "r") as f:
                        existing_messages = json.load(f)
                except (json.JSONDecodeError, IOError):
                    existing_messages = []

                # Merge messages (avoid duplicates by ID)
                existing_ids = {m.get("id") for m in existing_messages if m.get("id")}
                new_messages = [m for m in messages if m.get("id") not in existing_ids]
                messages = existing_messages + new_messages
                messages.sort(key=lambda m: m.get("timestamp") or "")
            elif fmt == "csv":
                # For CSV, append new messages
                # This is more complex, so for now just overwrite
                pass

    # Download attachments if requested
    url_mapping = {}
    if download_attachments and fmt in ("jira", "html", "docx") and messages:
        attachments_dir_name = output_path.stem + "_files"
        attachments_dir = output_path.parent / attachments_dir_name
        # Use parallel downloads for better performance
        if download_all_types:
            # Download all attachment types (PDFs, docs, etc.)
            url_mapping = _download_all_attachments_parallel(client, messages, attachments_dir, max_workers=5)
        else:
            # Download only images (backward compatible)
            url_mapping = _download_attachments_parallel(client, messages, attachments_dir, max_workers=5)

    # Write output
    if fmt == "json":
        _write_json(messages, output_path)
    elif fmt == "csv":
        _write_csv(messages, output_path)
    elif fmt == "jira":
        chat_title = chat.get("topic") or chat.get("displayName") or identifier
        participants_list = _member_labels(chat)
        chat_info = {
            "title": chat_title,
            "participants": ", ".join(participants_list) if participants_list else "N/A",
            "date_range": f"{start_dt.date()} to {end_dt.date()}",
        }
        write_jira_markdown(messages, output_path, chat_info=chat_info, url_mapping=url_mapping)
    elif fmt == "html":
        chat_title = chat.get("topic") or chat.get("displayName") or identifier
        participants_list = _member_labels(chat)
        chat_info = {
            "title": chat_title,
            "participants": ", ".join(participants_list) if participants_list else "N/A",
            "date_range": f"{start_dt.date()} to {end_dt.date()}",
        }
        write_html(messages, output_path, chat_info=chat_info, url_mapping=url_mapping)
    elif fmt == "docx":
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

    return output_path, message_count, True
