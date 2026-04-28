"""Tests for export functionality."""

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import MagicMock, patch, call

import pytest

from teams_export.exporter import (
    choose_chat,
    export_chat,
    export_chat_incremental,
    ChatNotFoundError,
    _transform_message,
    _extract_attachment_urls,
    _normalise_filename,
    _within_range,
)
from teams_export.delta import DeltaStateManager


class TestChooseChat:
    """Test chat selection logic."""

    def test_choose_by_participant_email(self, sample_chat):
        """Test selecting chat by participant email."""
        chats = [sample_chat]
        result = choose_chat(chats, participant="alice@example.com")
        assert result["id"] == sample_chat["id"]

    def test_choose_by_participant_name(self, sample_chat):
        """Test selecting chat by participant display name."""
        chats = [sample_chat]
        result = choose_chat(chats, participant="Alice Smith")
        assert result["id"] == sample_chat["id"]

    def test_choose_by_chat_name(self, sample_chat):
        """Test selecting chat by chat display name."""
        chats = [sample_chat]
        result = choose_chat(chats, chat_name="Test Chat")
        assert result["id"] == sample_chat["id"]

    def test_choose_multiple_matches(self, sample_chat):
        """Test when multiple chats match criteria."""
        chat2 = sample_chat.copy()
        chat2["id"] = "different_id"
        chats = [sample_chat, chat2]

        result = choose_chat(chats, chat_name="Test Chat")
        assert isinstance(result, list)
        assert len(result) == 2

    def test_choose_no_matches(self, sample_chat):
        """Test when no chats match criteria."""
        chats = [sample_chat]
        with pytest.raises(ChatNotFoundError):
            choose_chat(chats, participant="nonexistent@example.com")

    def test_choose_case_insensitive(self, sample_chat):
        """Test case-insensitive matching."""
        chats = [sample_chat]
        result = choose_chat(chats, participant="ALICE@EXAMPLE.COM")
        assert result["id"] == sample_chat["id"]


class TestTransformMessage:
    """Test message transformation."""

    def test_transform_basic_message(self):
        """Test transforming a basic message."""
        message = {
            "id": "msg123",
            "createdDateTime": "2025-01-15T10:00:00Z",
            "lastModifiedDateTime": "2025-01-15T10:00:00Z",
            "from": {
                "user": {
                    "displayName": "John Doe",
                    "userPrincipalName": "john@example.com"
                }
            },
            "body": {
                "content": "Hello world",
                "contentType": "text"
            },
            "messageType": "message"
        }

        transformed = _transform_message(message)
        assert transformed["id"] == "msg123"
        assert transformed["sender"] == "John Doe"
        assert transformed["sender_email"] == "john@example.com"
        assert transformed["content"] == "Hello world"
        assert transformed["content_type"] == "text"
        assert transformed["type"] == "message"

    def test_transform_message_with_attachments(self):
        """Test transforming a message with attachments."""
        message = {
            "id": "msg456",
            "from": {
                "user": {
                    "displayName": "Jane Doe"
                }
            },
            "body": {
                "content": "See attached",
                "contentType": "text"
            },
            "attachments": [
                {
                    "id": "att1",
                    "name": "document.pdf",
                    "contentType": "application/pdf"
                }
            ],
            "reactions": [
                {"reactionType": "like"}
            ]
        }

        transformed = _transform_message(message)
        assert len(transformed["attachments"]) == 1
        assert transformed["attachments"][0]["name"] == "document.pdf"
        assert len(transformed["reactions"]) == 1


class TestExtractAttachments:
    """Test attachment extraction."""

    def test_extract_image_urls_from_html(self):
        """Test extracting image URLs from HTML content."""
        messages = [
            {
                "content": '<p>Check this <img src="https://example.com/img1.png"> and <img src="https://example.com/img2.jpg"></p>'
            }
        ]

        attachments = _extract_attachment_urls(messages, images_only=True)
        urls = [url for url, _, _ in attachments]
        assert "https://example.com/img1.png" in urls
        assert "https://example.com/img2.jpg" in urls

    def test_extract_all_attachments(self):
        """Test extracting all attachment types."""
        messages = [
            {
                "attachments": [
                    {
                        "name": "document.pdf",
                        "contentType": "application/pdf",
                        "contentUrl": "https://example.com/doc.pdf"
                    },
                    {
                        "name": "image.png",
                        "contentType": "image/png",
                        "contentUrl": "https://example.com/img.png"
                    },
                    {
                        "name": "spreadsheet.xlsx",
                        "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "contentUrl": "https://example.com/sheet.xlsx"
                    }
                ]
            }
        ]

        # Extract all attachments
        all_attachments = _extract_attachment_urls(messages, images_only=False)
        assert len(all_attachments) == 3

        # Extract only images
        image_attachments = _extract_attachment_urls(messages, images_only=True)
        assert len(image_attachments) == 1
        assert image_attachments[0][1] == "image.png"

    def test_extract_no_duplicates(self):
        """Test that duplicate URLs are not extracted."""
        messages = [
            {
                "content": '<img src="https://example.com/img.png">',
                "attachments": [
                    {
                        "contentUrl": "https://example.com/img.png",
                        "contentType": "image/png"
                    }
                ]
            },
            {
                "content": '<img src="https://example.com/img.png">'
            }
        ]

        attachments = _extract_attachment_urls(messages, images_only=True)
        # Should only have one entry despite appearing multiple times
        assert len(attachments) == 1


class TestFilenameNormalization:
    """Test filename normalization."""

    def test_normalise_filename_basic(self):
        """Test basic filename normalization."""
        assert _normalise_filename("Test Chat") == "test_chat"
        assert _normalise_filename("User@Example.com") == "user_example_com"
        assert _normalise_filename("Meeting: Project Review") == "meeting_project_review"

    def test_normalise_filename_special_chars(self):
        """Test normalization with special characters."""
        assert _normalise_filename("Test/Chat\\Name") == "test_chat_name"
        assert _normalise_filename("Chat #123") == "chat_123"
        assert _normalise_filename("  Spaces  ") == "spaces"


class TestDateRangeFiltering:
    """Test message date range filtering."""

    def test_within_range_inclusive(self):
        """Test that range filtering is inclusive."""
        now = datetime.now(timezone.utc)
        start = now - timedelta(hours=2)
        end = now

        message = {
            "createdDateTime": (now - timedelta(hours=1)).isoformat()
        }

        assert _within_range(message, start, end)

    def test_within_range_outside(self):
        """Test messages outside range are excluded."""
        now = datetime.now(timezone.utc)
        start = now - timedelta(hours=2)
        end = now - timedelta(hours=1)

        message = {
            "createdDateTime": now.isoformat()
        }

        assert not _within_range(message, start, end)

    def test_within_range_missing_timestamp(self):
        """Test messages without timestamp are excluded."""
        now = datetime.now(timezone.utc)
        message = {"id": "msg123"}  # No timestamp

        assert not _within_range(message, now, now)


class TestExportChat:
    """Test chat export functionality."""

    @patch("teams_export.exporter._download_attachments_parallel")
    def test_export_chat_json(self, mock_download, mock_graph_client, sample_chat, sample_messages, temp_dir):
        """Test exporting chat to JSON format."""
        mock_graph_client.list_chat_messages_parallel.return_value = sample_messages
        mock_download.return_value = {}

        now = datetime.now(timezone.utc)
        start = now - timedelta(days=1)
        end = now

        output_path, count = export_chat(
            mock_graph_client,
            sample_chat,
            start,
            end,
            output_dir=temp_dir,
            output_format="json",
            download_attachments=False
        )

        assert output_path.exists()
        assert output_path.suffix == ".json"
        assert count == 3  # All 3 sample messages

        # Check JSON content
        with open(output_path) as f:
            data = json.load(f)
            assert len(data) == 3
            assert data[0]["sender"] == "Alice Smith"

    @patch("teams_export.exporter._download_attachments_parallel")
    def test_export_chat_csv(self, mock_download, mock_graph_client, sample_chat, sample_messages, temp_dir):
        """Test exporting chat to CSV format."""
        mock_graph_client.list_chat_messages_parallel.return_value = sample_messages
        mock_download.return_value = {}

        now = datetime.now(timezone.utc)
        output_path, count = export_chat(
            mock_graph_client,
            sample_chat,
            now - timedelta(days=1),
            now,
            output_dir=temp_dir,
            output_format="csv"
        )

        assert output_path.exists()
        assert output_path.suffix == ".csv"
        assert count == 3

    @patch("teams_export.exporter._download_all_attachments_parallel")
    def test_export_with_all_attachments(self, mock_download, mock_graph_client, sample_chat, sample_messages, temp_dir):
        """Test exporting with all attachments download."""
        mock_graph_client.list_chat_messages_parallel.return_value = sample_messages
        mock_download.return_value = {
            "https://example.com/document.pdf": "Test_Chat_files/document.pdf"
        }

        now = datetime.now(timezone.utc)
        output_path, count = export_chat(
            mock_graph_client,
            sample_chat,
            now - timedelta(days=1),
            now,
            output_dir=temp_dir,
            output_format="jira",
            download_attachments=True,
            download_all_types=True
        )

        # Should have called the all attachments function
        mock_download.assert_called_once()


class TestIncrementalExport:
    """Test incremental export functionality."""

    @patch("teams_export.exporter._download_attachments_parallel")
    def test_incremental_first_sync(self, mock_download, mock_graph_client, sample_chat, sample_messages, temp_dir):
        """Test first incremental sync (full sync)."""
        mock_graph_client.list_chat_messages_delta.return_value = (
            sample_messages,
            "https://graph.microsoft.com/v1.0/delta?token=new_token"
        )
        mock_download.return_value = {}

        delta_manager = DeltaStateManager(temp_dir / "delta")
        now = datetime.now(timezone.utc)

        output_path, count, has_changes = export_chat_incremental(
            mock_graph_client,
            sample_chat,
            now - timedelta(days=1),
            now,
            output_dir=temp_dir,
            output_format="json",
            delta_manager=delta_manager
        )

        assert output_path.exists()
        assert count == 3
        assert has_changes

        # Check that delta state was saved
        state = delta_manager.get_state(sample_chat["id"])
        assert state is not None
        assert state["delta_link"] == "https://graph.microsoft.com/v1.0/delta?token=new_token"

    @patch("teams_export.exporter._download_attachments_parallel")
    def test_incremental_no_changes(self, mock_download, mock_graph_client, sample_chat, temp_dir):
        """Test incremental sync with no new messages."""
        # Return empty messages for delta sync
        mock_graph_client.list_chat_messages_delta.return_value = (
            [],
            "https://graph.microsoft.com/v1.0/delta?token=new_token"
        )
        mock_download.return_value = {}

        delta_manager = DeltaStateManager(temp_dir / "delta")
        # Save initial state
        delta_manager.save_state(
            sample_chat["id"],
            delta_link="https://graph.microsoft.com/v1.0/delta?token=old_token"
        )

        now = datetime.now(timezone.utc)
        output_path, count, has_changes = export_chat_incremental(
            mock_graph_client,
            sample_chat,
            now - timedelta(days=1),
            now,
            output_dir=temp_dir,
            output_format="json",
            delta_manager=delta_manager
        )

        assert not has_changes
        assert count == 0
