"""Local caching for chat lists to speed up repeated operations."""

from __future__ import annotations

import json
import time
from pathlib import Path
from typing import List, Optional


DEFAULT_CACHE_DIR = Path("~/.teams-exporter/cache").expanduser()
CACHE_TTL_SECONDS = 86400  # 24 hours (1 day)


class ChatCache:
    """Simple file-based cache for chat lists."""

    def __init__(self, cache_dir: Path = DEFAULT_CACHE_DIR):
        self.cache_dir = cache_dir
        self.cache_file = cache_dir / "chats_cache.json"

    def get(self, user_id: str) -> Optional[List[dict]]:
        """Get cached chats for a user if still valid.

        Args:
            user_id: User identifier (from token claims or 'me')

        Returns:
            List of chats if cache is valid, None otherwise
        """
        if not self.cache_file.exists():
            return None

        try:
            with self.cache_file.open("r", encoding="utf-8") as f:
                cache_data = json.load(f)

            # Check if cache is for the same user
            if cache_data.get("user_id") != user_id:
                return None

            # Check if cache is still fresh
            cached_time = cache_data.get("timestamp", 0)
            age = time.time() - cached_time
            if age > CACHE_TTL_SECONDS:
                return None

            chats = cache_data.get("chats", [])
            return chats if chats else None

        except (json.JSONDecodeError, KeyError, OSError):
            return None

    def set(self, user_id: str, chats: List[dict]) -> None:
        """Cache chat list for a user.

        Args:
            user_id: User identifier
            chats: List of chat objects to cache
        """
        self.cache_dir.mkdir(parents=True, exist_ok=True)

        cache_data = {
            "user_id": user_id,
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
