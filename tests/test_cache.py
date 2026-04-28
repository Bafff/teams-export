"""Tests for chat list caching."""

import json
import time
from datetime import datetime
from pathlib import Path

import pytest
from freezegun import freeze_time

from teams_export.cache import ChatCache


class TestChatCache:
    """Test chat list caching functionality."""

    def test_cache_init_creates_directory(self, tmp_path, monkeypatch):
        """Test that cache initialization creates the cache directory."""
        monkeypatch.setattr(Path, "home", lambda: tmp_path)
        cache_dir = tmp_path / ".teams-exporter" / "cache"
        assert not cache_dir.exists()

        cache = ChatCache()
        # Directory should be created on first operation
        cache.get("test_user")
        assert cache_dir.exists()

    def test_cache_set_and_get(self, tmp_path, monkeypatch):
        """Test setting and getting cache."""
        monkeypatch.setattr(Path, "home", lambda: tmp_path)
        cache = ChatCache()

        chats = [
            {"id": "chat1", "topic": "Test Chat 1"},
            {"id": "chat2", "topic": "Test Chat 2"}
        ]

        # Set cache
        cache.set("user123", chats)

        # Get cache
        cached_chats = cache.get("user123")
        assert cached_chats == chats

    def test_cache_expiry(self, tmp_path, monkeypatch):
        """Test that cache expires after TTL."""
        monkeypatch.setattr(Path, "home", lambda: tmp_path)
        cache = ChatCache()

        chats = [{"id": "chat1", "topic": "Test Chat"}]

        # Set cache
        with freeze_time("2025-01-15 10:00:00"):
            cache.set("user123", chats)
            # Should be valid immediately
            assert cache.get("user123") == chats

        # Check just before expiry (24 hours - 1 second)
        with freeze_time("2025-01-16 09:59:59"):
            assert cache.get("user123") == chats

        # Check after expiry (24 hours + 1 second)
        with freeze_time("2025-01-16 10:00:01"):
            assert cache.get("user123") is None

    def test_cache_different_users(self, tmp_path, monkeypatch):
        """Test that cache is user-specific."""
        monkeypatch.setattr(Path, "home", lambda: tmp_path)
        cache = ChatCache()

        chats1 = [{"id": "chat1", "topic": "User 1 Chat"}]
        chats2 = [{"id": "chat2", "topic": "User 2 Chat"}]

        cache.set("user1", chats1)
        cache.set("user2", chats2)

        assert cache.get("user1") == chats1
        assert cache.get("user2") == chats2

    def test_cache_clear(self, tmp_path, monkeypatch):
        """Test clearing cache."""
        monkeypatch.setattr(Path, "home", lambda: tmp_path)
        cache = ChatCache()

        chats = [{"id": "chat1", "topic": "Test Chat"}]
        cache.set("user123", chats)
        assert cache.get("user123") == chats

        # Clear cache
        cache.clear()
        assert cache.get("user123") is None

        # Cache file should be deleted
        cache_file = tmp_path / ".teams-exporter" / "cache" / "chats_cache.json"
        assert not cache_file.exists()

    def test_cache_corrupted_file(self, tmp_path, monkeypatch):
        """Test handling of corrupted cache file."""
        monkeypatch.setattr(Path, "home", lambda: tmp_path)
        cache_dir = tmp_path / ".teams-exporter" / "cache"
        cache_dir.mkdir(parents=True)
        cache_file = cache_dir / "chats_cache.json"

        # Write corrupted data
        cache_file.write_text("not valid json {")

        cache = ChatCache()
        # Should return None for corrupted cache
        assert cache.get("user123") is None

        # Should be able to set new cache
        chats = [{"id": "chat1"}]
        cache.set("user123", chats)
        assert cache.get("user123") == chats

    def test_cache_file_structure(self, tmp_path, monkeypatch):
        """Test the structure of the cache file."""
        monkeypatch.setattr(Path, "home", lambda: tmp_path)
        cache = ChatCache()

        chats = [{"id": "chat1", "topic": "Test Chat"}]
        cache.set("user123", chats)

        # Read cache file directly
        cache_file = tmp_path / ".teams-exporter" / "cache" / "chats_cache.json"
        with open(cache_file) as f:
            data = json.load(f)

        assert "user123" in data
        assert "timestamp" in data["user123"]
        assert "chats" in data["user123"]
        assert data["user123"]["chats"] == chats

    def test_cache_no_home_directory(self, monkeypatch):
        """Test cache behavior when home directory doesn't exist."""
        # Simulate no home directory
        monkeypatch.setattr(Path, "home", lambda: Path("/nonexistent/path"))

        cache = ChatCache()
        # Should handle gracefully
        assert cache.get("user123") is None

        # Set should also fail gracefully
        cache.set("user123", [{"id": "chat1"}])
        assert cache.get("user123") is None