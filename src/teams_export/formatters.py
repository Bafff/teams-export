"""Output formatters for Teams chat exports."""

from __future__ import annotations

import html
import re
import base64
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


def _format_jira_message(message: dict, index: int, url_mapping: dict[str, str] | None = None) -> str:
    """Format a single message in standard Markdown.

    Args:
        message: Message dictionary
        index: Message index
        url_mapping: Optional mapping of remote URL to local file path
    """
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
            # Use local path if available, otherwise use remote URL
            display_url = url_mapping.get(src, src) if url_mapping else src
            attachment_lines.append(f"![{alt}]({display_url})")

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
                # Use local path if available, otherwise use remote URL
                display_url = url_mapping.get(url, url) if url_mapping else url
                # Format as markdown image
                attachment_lines.append(f"![{name}]({display_url})")
            elif url:
                # Use local path if available, otherwise use remote URL
                display_url = url_mapping.get(url, url) if url_mapping else url
                # Format as markdown link
                attachment_lines.append(f"📎 [{name}]({display_url})")
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


def write_jira_markdown(
    messages: Sequence[dict],
    output_path: Path,
    chat_info: dict | None = None,
    url_mapping: dict[str, str] | None = None,
) -> None:
    """Write messages in standard Markdown format (works in Jira, GitHub, and other platforms).

    Args:
        messages: List of message dictionaries
        output_path: Path to write markdown file
        chat_info: Optional chat metadata (title, participants, date range)
        url_mapping: Optional mapping of remote URLs to local file paths
    """

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
            lines.append(_format_jira_message(message, idx, url_mapping=url_mapping))
            # No extra empty line needed - _format_jira_message adds it
    else:
        lines.append("*No messages found in the specified date range.*")

    # Write to file
    content = "\n".join(lines)
    output_path.write_text(content, encoding="utf-8")


def _image_to_base64(image_path: Path) -> str | None:
    """Convert image file to base64 data URL.

    Returns:
        Data URL string like "data:image/png;base64,iVBORw0KG..." or None if failed
    """
    try:
        # Read image bytes
        image_bytes = image_path.read_bytes()

        # Encode to base64
        base64_data = base64.b64encode(image_bytes).decode('utf-8')

        # Determine MIME type from extension
        ext = image_path.suffix.lower()
        mime_type = {
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.bmp': 'image/bmp',
            '.webp': 'image/webp',
            '.svg': 'image/svg+xml',
            '.tiff': 'image/tiff',
            '.tif': 'image/tiff',
        }.get(ext, 'image/png')

        return f"data:{mime_type};base64,{base64_data}"
    except Exception as e:
        print(f"Warning: Failed to encode {image_path}: {e}")
        return None


def _format_html_message(message: dict, index: int, url_mapping: dict[str, str] | None = None, base_dir: Path | None = None) -> str:
    """Format a single message as HTML with embedded images.

    Args:
        message: Message dictionary
        index: Message index
        url_mapping: Mapping of remote URLs to local file paths
        base_dir: Base directory for resolving relative image paths
    """
    sender = message.get("sender") or "Unknown"
    timestamp = message.get("timestamp", "")

    # Format timestamp
    if timestamp:
        try:
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

    # Strip HTML to get text content
    content = _strip_html(html_content)

    # Escape HTML in content
    content = html.escape(content) if content else ""

    # Replace newlines with <br>
    content = content.replace('\n', '<br>')

    # Format attachments
    attachments = message.get("attachments", [])
    attachment_html = []

    # Add inline images first
    for img in inline_images:
        src = img.get("src", "")
        alt = img.get("alt", "image")
        if src:
            # Try to get local path from url_mapping
            local_path = url_mapping.get(src) if url_mapping else None

            if local_path and base_dir:
                # Convert local file to base64
                try:
                    img_path = base_dir / local_path
                    if img_path.exists():
                        data_url = _image_to_base64(img_path)
                        if data_url:
                            src = data_url
                except Exception:
                    pass  # Keep original URL if conversion fails

            attachment_html.append(f'<img src="{html.escape(src)}" alt="{html.escape(alt)}" style="max-width: 100%; height: auto; margin: 10px 0;">')

    # Then add file attachments
    if attachments:
        for att in attachments:
            name = att.get("name") or "Attachment"
            content_type = att.get("contentType", "")

            url = (
                att.get("contentUrl") or
                att.get("content") or
                att.get("url") or
                att.get("thumbnailUrl") or
                (att.get("hostedContents", {}).get("contentUrl") if isinstance(att.get("hostedContents"), dict) else None)
            )

            is_image = (
                content_type.startswith("image/") if content_type else
                any(name.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp'])
            )

            if is_image and url:
                # Try to get local path from url_mapping
                local_path = url_mapping.get(url) if url_mapping else None

                if local_path and base_dir:
                    # Convert local file to base64
                    try:
                        img_path = base_dir / local_path
                        if img_path.exists():
                            data_url = _image_to_base64(img_path)
                            if data_url:
                                url = data_url
                    except Exception:
                        pass  # Keep original URL if conversion fails

                attachment_html.append(f'<img src="{html.escape(url)}" alt="{html.escape(name)}" style="max-width: 100%; height: auto; margin: 10px 0;">')
            elif url:
                # Try to get local path from url_mapping for non-image attachments
                local_path = url_mapping.get(url) if url_mapping else None
                display_url = local_path if local_path else url
                attachment_html.append(f'<p>📎 <a href="{html.escape(display_url)}">{html.escape(name)}</a></p>')
            else:
                attachment_html.append(f'<p>📎 {html.escape(name)} (no URL)</p>')

    # Handle empty content
    if not content:
        content_type = message.get("type", "")
        if content_type == "systemEventMessage":
            content = "[System event]"
        elif not attachment_html:
            content = "[No content]"

    # Format reactions
    reactions = message.get("reactions", [])
    reaction_html = ""
    if reactions:
        reaction_emojis = []
        for reaction in reactions:
            reaction_type = reaction.get("reactionType", "")
            if reaction_type:
                reaction_emojis.append(html.escape(reaction_type))
        if reaction_emojis:
            reaction_html = f" <span style='color: #666;'>[{', '.join(reaction_emojis)}]</span>"

    # Build HTML message block
    html_parts = [
        f'<div style="margin-bottom: 20px; padding: 15px; border-left: 3px solid #0078d4; background-color: #f5f5f5;">',
        f'<div style="font-weight: bold; color: #333;">{html.escape(sender)} <span style="color: #666; font-weight: normal;">— {timestamp_clean}</span>{reaction_html}</div>',
    ]

    if content:
        html_parts.append(f'<div style="margin-top: 10px; color: #555; padding-left: 15px; border-left: 2px solid #ddd;">{content}</div>')

    if attachment_html:
        html_parts.append('<div style="margin-top: 10px;">')
        html_parts.extend(attachment_html)
        html_parts.append('</div>')

    html_parts.append('</div>')

    return "\n".join(html_parts)


def write_html(
    messages: Sequence[dict],
    output_path: Path,
    chat_info: dict | None = None,
    url_mapping: dict[str, str] | None = None,
) -> None:
    """Write messages as HTML with embedded base64 images.

    This format is perfect for copy-pasting into Jira/Confluence:
    1. Open the HTML file in a browser
    2. Select all (Ctrl+A)
    3. Copy (Ctrl+C)
    4. Paste into Jira/Confluence - images will be embedded!

    Args:
        messages: List of message dictionaries
        output_path: Path to write HTML file
        chat_info: Optional chat metadata (title, participants, date range)
        url_mapping: Optional mapping of remote URLs to local file paths
    """
    html_parts = [
        '<!DOCTYPE html>',
        '<html>',
        '<head>',
        '    <meta charset="utf-8">',
        '    <meta name="viewport" content="width=device-width, initial-scale=1.0">',
        '    <title>Teams Chat Export</title>',
        '    <style>',
        '        body {',
        '            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;',
        '            line-height: 1.6;',
        '            max-width: 900px;',
        '            margin: 0 auto;',
        '            padding: 20px;',
        '            background-color: #fff;',
        '            color: #333;',
        '        }',
        '        h2 { color: #0078d4; border-bottom: 2px solid #0078d4; padding-bottom: 10px; }',
        '        .meta { color: #666; margin-bottom: 20px; }',
        '        .meta strong { color: #333; }',
        '        hr { border: none; border-top: 1px solid #ddd; margin: 30px 0; }',
        '        img { display: block; }',
        '        .copy-btn {',
        '            position: fixed;',
        '            top: 20px;',
        '            right: 20px;',
        '            padding: 10px 20px;',
        '            background-color: #0078d4;',
        '            color: white;',
        '            border: none;',
        '            border-radius: 4px;',
        '            cursor: pointer;',
        '            font-size: 14px;',
        '            box-shadow: 0 2px 8px rgba(0,0,0,0.2);',
        '            z-index: 1000;',
        '        }',
        '        .copy-btn:hover { background-color: #005a9e; }',
        '        .copy-btn:active { background-color: #004578; }',
        '        .copy-status {',
        '            position: fixed;',
        '            top: 60px;',
        '            right: 20px;',
        '            padding: 8px 16px;',
        '            background-color: #107c10;',
        '            color: white;',
        '            border-radius: 4px;',
        '            font-size: 13px;',
        '            display: none;',
        '            z-index: 1001;',
        '        }',
        '    </style>',
        '    <script>',
        '        async function base64ToBlob(base64, mimeType) {',
        '            const response = await fetch(base64);',
        '            return response.blob();',
        '        }',
        '',
        '        async function copyContent() {',
        '            const button = document.getElementById("copyBtn");',
        '            const status = document.getElementById("copyStatus");',
        '            ',
        '            try {',
        '                button.disabled = true;',
        '                button.textContent = "Copying...";',
        '',
        '                // Get the content div',
        '                const content = document.getElementById("content");',
        '                ',
        '                // Clone the content to avoid modifying the original',
        '                const clone = content.cloneNode(true);',
        '                ',
        '                // Find all images with base64 data',
        '                const images = clone.querySelectorAll("img[src^=\'data:\']");',
        '                const imageBlobs = [];',
        '                ',
        '                // Convert base64 images to blob URLs for better clipboard support',
        '                for (let img of images) {',
        '                    try {',
        '                        const blob = await base64ToBlob(img.src, "image/png");',
        '                        const blobUrl = URL.createObjectURL(blob);',
        '                        img.src = blobUrl;',
        '                        imageBlobs.push({ img, blob, blobUrl });',
        '                    } catch (e) {',
        '                        console.error("Failed to convert image:", e);',
        '                    }',
        '                }',
        '',
        '                // Create HTML blob for clipboard',
        '                const htmlBlob = new Blob([clone.outerHTML], { type: "text/html" });',
        '                const textBlob = new Blob([clone.textContent], { type: "text/plain" });',
        '',
        '                // Write to clipboard',
        '                await navigator.clipboard.write([',
        '                    new ClipboardItem({',
        '                        "text/html": htmlBlob,',
        '                        "text/plain": textBlob',
        '                    })',
        '                ]);',
        '',
        '                // Clean up blob URLs',
        '                imageBlobs.forEach(({ blobUrl }) => URL.revokeObjectURL(blobUrl));',
        '',
        '                // Show success message',
        '                status.textContent = "✓ Copied! Now paste into Jira/Confluence";',
        '                status.style.display = "block";',
        '                setTimeout(() => { status.style.display = "none"; }, 3000);',
        '                ',
        '            } catch (err) {',
        '                console.error("Copy failed:", err);',
        '                status.textContent = "❌ Copy failed. Try selecting manually (Ctrl+A)";',
        '                status.style.backgroundColor = "#d13438";',
        '                status.style.display = "block";',
        '                setTimeout(() => { status.style.display = "none"; }, 3000);',
        '            } finally {',
        '                button.disabled = false;',
        '                button.textContent = "📋 Copy to Clipboard";',
        '            }',
        '        }',
        '',
        '        // Alternative: handle manual copy (Ctrl+C) on selected content',
        '        document.addEventListener("copy", async (e) => {',
        '            const selection = window.getSelection();',
        '            if (!selection.rangeCount) return;',
        '',
        '            const container = document.createElement("div");',
        '            container.appendChild(selection.getRangeAt(0).cloneContents());',
        '            ',
        '            const images = container.querySelectorAll("img[src^=\'data:\']");',
        '            if (images.length > 0) {',
        '                e.preventDefault();',
        '                ',
        '                // Convert base64 images to blob URLs',
        '                for (let img of images) {',
        '                    try {',
        '                        const blob = await base64ToBlob(img.src, "image/png");',
        '                        const blobUrl = URL.createObjectURL(blob);',
        '                        img.src = blobUrl;',
        '                    } catch (err) {',
        '                        console.error("Failed to convert image:", err);',
        '                    }',
        '                }',
        '',
        '                // Set clipboard data',
        '                e.clipboardData.setData("text/html", container.innerHTML);',
        '                e.clipboardData.setData("text/plain", container.textContent);',
        '            }',
        '        });',
        '    </script>',
        '</head>',
        '<body>',
        '    <button id="copyBtn" class="copy-btn" onclick="copyContent()">📋 Copy to Clipboard</button>',
        '    <div id="copyStatus" class="copy-status"></div>',
        '    <div id="content">',
    ]

    # Add header with chat info
    if chat_info:
        chat_title = chat_info.get("title", "Teams Chat Export")
        participants = chat_info.get("participants", "")
        date_range = chat_info.get("date_range", "")

        html_parts.append(f'<h2>{html.escape(chat_title)}</h2>')
        if participants:
            html_parts.append(f'<p class="meta"><strong>Participants:</strong> {html.escape(participants)}</p>')
        if date_range:
            html_parts.append(f'<p class="meta"><strong>Date Range:</strong> {html.escape(date_range)}</p>')
        html_parts.append('<hr>')

    # Add messages
    if messages:
        html_parts.append(f'<h3>Messages ({len(messages)} total)</h3>')

        for idx, message in enumerate(messages, 1):
            html_parts.append(_format_html_message(message, idx, url_mapping=url_mapping, base_dir=output_path.parent))
    else:
        html_parts.append('<p><em>No messages found in the specified date range.</em></p>')

    html_parts.extend([
        '    </div>',  # Close content div
        '</body>',
        '</html>',
    ])

    # Write to file
    content = "\n".join(html_parts)
    output_path.write_text(content, encoding="utf-8")
