"""pytest configuration and shared fixtures."""

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import MagicMock

import pytest


@pytest.fixture
def sample_chat():
    """Sample chat object."""
    return {
        "id": "19:meeting_test@thread.v2",
        "topic": "Test Chat",
        "displayName": "Test Chat",
        "chatType": "oneOnOne",
        "members": [
            {
                "id": "user1",
                "displayName": "Alice Smith",
                "email": "alice@example.com",
                "userPrincipalName": "alice@example.com"
            },
            {
                "id": "user2",
                "displayName": "Bob Jones",
                "email": "bob@example.com",
                "userPrincipalName": "bob@example.com"
            }
        ],
        "lastMessagePreview": {
            "id": "msg123",
            "createdDateTime": datetime.now(timezone.utc).isoformat(),
            "body": {
                "content": "Last message",
                "contentType": "text"
            }
        }
    }


@pytest.fixture
def sample_messages():
    """Sample message list."""
    now = datetime.now(timezone.utc)
    return [
        {
            "id": "msg1",
            "createdDateTime": (now - timedelta(hours=2)).isoformat(),
            "lastModifiedDateTime": (now - timedelta(hours=2)).isoformat(),
            "from": {
                "user": {
                    "displayName": "Alice Smith",
                    "userPrincipalName": "alice@example.com"
                }
            },
            "body": {
                "content": "Hello everyone!",
                "contentType": "text"
            },
            "messageType": "message"
        },
        {
            "id": "msg2",
            "createdDateTime": (now - timedelta(hours=1)).isoformat(),
            "lastModifiedDateTime": (now - timedelta(hours=1)).isoformat(),
            "from": {
                "user": {
                    "displayName": "Bob Jones",
                    "userPrincipalName": "bob@example.com"
                }
            },
            "body": {
                "content": '<p>Hi Alice! Check out this <img src="https://example.com/image.png" alt="image">.</p>',
                "contentType": "html"
            },
            "messageType": "message",
            "attachments": [
                {
                    "id": "att1",
                    "name": "document.pdf",
                    "contentType": "application/pdf",
                    "contentUrl": "https://example.com/document.pdf"
                }
            ]
        },
        {
            "id": "msg3",
            "createdDateTime": now.isoformat(),
            "lastModifiedDateTime": now.isoformat(),
            "from": {
                "user": {
                    "displayName": "Alice Smith",
                    "userPrincipalName": "alice@example.com"
                }
            },
            "body": {
                "content": "Thanks Bob!",
                "contentType": "text"
            },
            "messageType": "message",
            "reactions": [
                {
                    "reactionType": "like",
                    "user": {
                        "displayName": "Bob Jones"
                    }
                }
            ]
        }
    ]


@pytest.fixture
def temp_dir(tmp_path):
    """Temporary directory for test outputs."""
    return tmp_path


@pytest.fixture
def mock_graph_client():
    """Mock GraphClient for testing."""
    client = MagicMock()
    client._session = MagicMock()
    client._base_url = "https://graph.microsoft.com/v1.0"
    return client


@pytest.fixture
def config_dir(tmp_path):
    """Temporary config directory."""
    config_path = tmp_path / ".teams-exporter"
    config_path.mkdir()
    return config_path


@pytest.fixture
def sample_config(config_dir):
    """Sample configuration."""
    config = {
        "client_id": "test-client-id",
        "authority": "https://login.microsoftonline.com/common",
        "scopes": ["Chat.Read", "Chat.ReadBasic"]
    }
    config_file = config_dir / "config.json"
    config_file.write_text(json.dumps(config))
    return config
