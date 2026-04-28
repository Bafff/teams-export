"""Local caching for chat lists to speed up repeated operations."""

from __future__ import annotations

import json
import time
from pathlib import Path
from typing import List, Optional


DEFAULT_CACHE_SUBDIR = Path(".teams-exporter/cache")
CACHE_TTL_SECONDS = 86400  # 24 hours (1 day)


class ChatCache:
    """Simple file-based cache for chat lists."""

    def __init__(self, cache_dir: Path | None = None):
        if cache_dir is None:
            cache_dir = Path.home() / DEFAULT_CACHE_SUBDIR
        else:
            cache_dir = cache_dir.expanduser()

        self.cache_dir = cache_dir
        self.cache_file = cache_dir / "chats_cache.json"
        self._enabled = self._ensure_directory()

    def _ensure_directory(self) -> bool:
        try:
            self.cache_dir.mkdir(parents=True, exist_ok=True)
            return True
        except OSError:
            return False

    def _load_cache_data(self) -> dict:
        if not self._enabled or not self.cache_file.exists():
            return {}

        try:
            with self.cache_file.open("r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (json.JSONDecodeError, OSError):
            return {}

    def get(self, user_id: str) -> Optional[List[dict]]:
        """Get cached chats for a user if still valid.

        Args:
            user_id: User identifier (from token claims or 'me')

        Returns:
            List of chats if cache is valid, None otherwise
        """
        cache_data = self._load_cache_data()
        entry = cache_data.get(user_id)
        if not entry:
            return None

        cached_time = entry.get("timestamp", 0)
        age = time.time() - cached_time
        if age > CACHE_TTL_SECONDS:
            return None

        chats = entry.get("chats") or []
        return chats or None

    def set(self, user_id: str, chats: List[dict]) -> None:
        """Cache chat list for a user.

        Args:
            user_id: User identifier
            chats: List of chat objects to cache
        """
        if not self._enabled:
            self._enabled = self._ensure_directory()
        if not self._enabled:
            return

        cache_data = self._load_cache_data()
        cache_data[user_id] = {
            "timestamp": time.time(),
            "chats": chats,
        }

        try:
            with self.cache_file.open("w", encoding="utf-8") as f:
                json.dump(cache_data, f, indent=2)
        except OSError:
            # Silently fail if can't write cache
            pass

    def clear(self) -> None:
        """Clear the cache file."""
        try:
            if self.cache_file.exists():
                self.cache_file.unlink()
        except OSError:
            pass
