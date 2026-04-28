"""Tests for delta sync state management."""

import json
from datetime import datetime, timezone
from pathlib import Path

import pytest

from teams_export.delta import DeltaStateManager


class TestDeltaStateManager:
    """Test delta state management functionality."""

    def test_init_creates_directory(self, tmp_path):
        """Test that initialization creates the state directory."""
        state_dir = tmp_path / "delta_states"
        assert not state_dir.exists()

        manager = DeltaStateManager(state_dir)
        assert state_dir.exists()
        assert state_dir.is_dir()

    def test_save_and_get_state(self, tmp_path):
        """Test saving and retrieving state."""
        manager = DeltaStateManager(tmp_path / "delta")
        chat_id = "19:test_chat@thread.v2"

        # Save state
        manager.save_state(
            chat_id,
            delta_link="https://graph.microsoft.com/v1.0/delta?token=abc123",
            last_message_id="msg_456",
            last_message_time="2025-01-15T10:00:00Z",
            message_count=42
        )

        # Retrieve state
        state = manager.get_state(chat_id)
        assert state is not None
        assert state["chat_id"] == chat_id
        assert state["delta_link"] == "https://graph.microsoft.com/v1.0/delta?token=abc123"
        assert state["last_message_id"] == "msg_456"
        assert state["last_message_time"] == "2025-01-15T10:00:00Z"
        assert state["message_count"] == 42
        assert "last_sync" in state

    def test_get_nonexistent_state(self, tmp_path):
        """Test getting state for a chat with no saved state."""
        manager = DeltaStateManager(tmp_path / "delta")
        state = manager.get_state("nonexistent_chat")
        assert state is None

    def test_clear_state(self, tmp_path):
        """Test clearing state for a specific chat."""
        manager = DeltaStateManager(tmp_path / "delta")
        chat_id = "19:test_chat@thread.v2"

        # Save state
        manager.save_state(chat_id, delta_link="test_link")
        assert manager.has_state(chat_id)

        # Clear state
        manager.clear_state(chat_id)
        assert not manager.has_state(chat_id)
        assert manager.get_state(chat_id) is None

    def test_clear_all_states(self, tmp_path):
        """Test clearing all saved states."""
        manager = DeltaStateManager(tmp_path / "delta")

        # Save multiple states
        manager.save_state("chat1", delta_link="link1")
        manager.save_state("chat2", delta_link="link2")
        manager.save_state("chat3", delta_link="link3")

        assert manager.has_state("chat1")
        assert manager.has_state("chat2")
        assert manager.has_state("chat3")

        # Clear all
        manager.clear_all_states()

        assert not manager.has_state("chat1")
        assert not manager.has_state("chat2")
        assert not manager.has_state("chat3")

    def test_list_states(self, tmp_path):
        """Test listing all saved states."""
        manager = DeltaStateManager(tmp_path / "delta")

        # Save multiple states
        manager.save_state("chat1", delta_link="link1", message_count=10)
        manager.save_state("chat2", delta_link="link2", message_count=20)
        manager.save_state("chat3", delta_link="link3", message_count=30)

        states = manager.list_states()
        assert len(states) == 3
        assert "chat1" in states
        assert "chat2" in states
        assert "chat3" in states
        assert states["chat1"]["message_count"] == 10
        assert states["chat2"]["message_count"] == 20
        assert states["chat3"]["message_count"] == 30

    def test_sanitize_chat_id(self, tmp_path):
        """Test that chat IDs are sanitized for filesystem safety."""
        manager = DeltaStateManager(tmp_path / "delta")

        # Chat ID with special characters
        chat_id = "19:meeting/with:slashes@thread.v2"
        manager.save_state(chat_id, delta_link="test")

        # Should still be retrievable with original ID
        state = manager.get_state(chat_id)
        assert state is not None
        assert state["chat_id"] == chat_id

        # Check that file was created with safe name
        state_files = list((tmp_path / "delta").glob("*.json"))
        assert len(state_files) == 1
        assert "/" not in state_files[0].name
        assert ":" not in state_files[0].name.replace("_", "")  # Colons should be replaced

    def test_corrupted_state_file(self, tmp_path):
        """Test handling of corrupted state files."""
        manager = DeltaStateManager(tmp_path / "delta")
        chat_id = "test_chat"

        # Create corrupted state file
        state_file = manager._get_state_file(chat_id)
        state_file.write_text("not valid json {")

        # Should return None and remove corrupted file
        state = manager.get_state(chat_id)
        assert state is None
        assert not state_file.exists()

    def test_update_existing_state(self, tmp_path):
        """Test updating an existing state."""
        manager = DeltaStateManager(tmp_path / "delta")
        chat_id = "test_chat"

        # Initial save
        manager.save_state(chat_id, delta_link="link1", message_count=10)
        state1 = manager.get_state(chat_id)
        sync_time1 = state1["last_sync"]

        # Update state
        manager.save_state(chat_id, delta_link="link2", message_count=20)
        state2 = manager.get_state(chat_id)

        assert state2["delta_link"] == "link2"
        assert state2["message_count"] == 20
        assert state2["last_sync"] != sync_time1  # Should have new timestamp