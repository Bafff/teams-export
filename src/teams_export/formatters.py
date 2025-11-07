"""Output formatters for Teams chat exports."""

from __future__ import annotations

import html
import re
from typing import Sequence
from pathlib import Path


def _extract_images_from_html(content: str | None) -> list[dict]:
    """Extract inline images from HTML content.

    Returns list of dicts with 'src' and 'alt' keys.
    """
    if not content:
        return []

    images = []
    # Find all <img> tags and extract src and alt attributes
    img_pattern = r'<img[^>]+src=["\']([^"\']+)["\'][^>]*>'
    for match in re.finditer(img_pattern, content, flags=re.IGNORECASE):
        img_tag = match.group(0)
        src = match.group(1)

        # Try to extract alt text
        alt_match = re.search(r'alt=["\']([^"\']*)["\']', img_tag, flags=re.IGNORECASE)
        alt = alt_match.group(1) if alt_match else "image"

        # Try to extract itemid for better name
        itemid_match = re.search(r'itemid=["\']([^"\']+)["\']', img_tag, flags=re.IGNORECASE)
        if itemid_match and itemid_match.group(1):
            alt = itemid_match.group(1)

        images.append({"src": src, "alt": alt})

    return images


def _strip_html(content: str | None) -> str:
    """Remove HTML tags and decode entities to plain text."""
    if not content:
        return ""

    # Decode HTML entities first
    text = html.unescape(content)

    # Remove <img> tags (they are extracted separately)
    text = re.sub(r'<img[^>]+>', '', text, flags=re.IGNORECASE)

    # Replace common HTML elements with markdown equivalents
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<p[^>]*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</p>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<div[^>]*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</div>', '\n', text, flags=re.IGNORECASE)

    # Bold and italic
    text = re.sub(r'<strong[^>]*>(.*?)</strong>', r'*\1*', text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<b[^>]*>(.*?)</b>', r'*\1*', text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<em[^>]*>(.*?)</em>', r'_\1_', text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<i[^>]*>(.*?)</i>', r'_\1_', text, flags=re.IGNORECASE | re.DOTALL)

    # Links - convert to [text](url) format
    text = re.sub(r'<a[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', r'\2 (\1)', text, flags=re.IGNORECASE | re.DOTALL)

    # Remove all other HTML tags
    text = re.sub(r'<[^>]+>', '', text)

    # Clean up excessive whitespace
    text = re.sub(r'\n\s*\n\s*\n+', '\n\n', text)
    text = text.strip()

    return text


def _format_jira_message(message: dict, index: int) -> str:
    """Format a single message in standard Markdown."""
    sender = message.get("sender") or "Unknown"
    timestamp = message.get("timestamp", "")

    # Format timestamp to be more readable
    if timestamp:
        # Extract just the date and time, skip milliseconds
        try:
            # Format: 2025-10-23T14:30:45.123Z -> 2025-10-23 14:30
            timestamp_clean = timestamp.split('.')[0].replace('T', ' ')
            if 'Z' in timestamp:
                timestamp_clean = timestamp_clean.replace('Z', ' UTC')
        except Exception:
            timestamp_clean = timestamp
    else:
        timestamp_clean = "No timestamp"

    # Extract inline images from HTML content first
    html_content = message.get("content", "")
    inline_images = _extract_images_from_html(html_content)

    # Then strip HTML to get text content
    content = _strip_html(html_content)

    # Format attachments if present
    attachments = message.get("attachments", [])
    attachment_lines = []

    # Add inline images first
    for img in inline_images:
        src = img.get("src", "")
        alt = img.get("alt", "image")
        if src:
            attachment_lines.append(f"![{alt}]({src})")

    # Then add file attachments
    if attachments:
        for att in attachments:
            name = att.get("name") or "Attachment"
            content_type = att.get("contentType", "")

            # Try to get URL from different possible fields (in order of preference)
            url = (
                att.get("contentUrl") or
                att.get("content") or
                att.get("url") or
                att.get("thumbnailUrl") or
                (att.get("hostedContents", {}).get("contentUrl") if isinstance(att.get("hostedContents"), dict) else None)
            )

            # Check if it's an image
            is_image = (
                content_type.startswith("image/") if content_type else
                any(name.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp'])
            )

            if is_image and url:
                # Format as markdown image
                attachment_lines.append(f"![{name}]({url})")
            elif url:
                # Format as markdown link
                attachment_lines.append(f"📎 [{name}]({url})")
            else:
                # Just show the name if no URL found
                attachment_lines.append(f"📎 {name} (no URL)")

    # Handle empty content
    if not content:
        content_type = message.get("type", "")
        if content_type == "systemEventMessage":
            content = "[System event]"
        elif not attachment_lines:
            # Only show "[No content]" if there are no attachments either
            content = "[No content]"

    # Format reactions if present
    reactions = message.get("reactions", [])
    reaction_text = ""
    if reactions:
        reaction_emojis = []
        for reaction in reactions:
            reaction_type = reaction.get("reactionType", "")
            if reaction_type:
                reaction_emojis.append(reaction_type)
        if reaction_emojis:
            reaction_text = f" [{', '.join(reaction_emojis)}]"

    # Build the message block in standard Markdown format
    lines = [
        f"**{sender}** — *{timestamp_clean}*{reaction_text}",
        "",
    ]

    # Add content if present
    if content:
        # Format content as blockquote (add '> ' prefix to each line)
        content_lines = content.split('\n')
        quoted_content = '\n'.join(f"> {line}" if line else ">" for line in content_lines)
        lines.append(quoted_content)
        lines.append("")

    # Add attachments if present
    if attachment_lines:
        lines.extend(attachment_lines)
        lines.append("")

    return "\n".join(lines)


def write_jira_markdown(messages: Sequence[dict], output_path: Path, chat_info: dict | None = None) -> None:
    """Write messages in standard Markdown format (works in Jira, GitHub, and other platforms)."""

    lines = []

    # Add header with chat info
    if chat_info:
        chat_title = chat_info.get("title", "Teams Chat Export")
        participants = chat_info.get("participants", "")
        date_range = chat_info.get("date_range", "")

        lines.append(f"## {chat_title}")
        lines.append("")
        if participants:
            lines.append(f"**Participants:** {participants}")
        if date_range:
            lines.append(f"**Date Range:** {date_range}")
        lines.append("")
        lines.append("---")
        lines.append("")

    # Add messages
    if messages:
        lines.append(f"### Messages ({len(messages)} total)")
        lines.append("")

        for idx, message in enumerate(messages, 1):
            lines.append(_format_jira_message(message, idx))
            # No extra empty line needed - _format_jira_message adds it
    else:
        lines.append("*No messages found in the specified date range.*")

    # Write to file
    content = "\n".join(lines)
    output_path.write_text(content, encoding="utf-8")
