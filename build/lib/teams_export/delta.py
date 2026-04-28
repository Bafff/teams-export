"""Delta sync state management for incremental exports."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, Optional
from datetime import datetime, timezone


class DeltaStateManager:
    """Manage delta tokens and sync state for incremental exports."""

    def __init__(self, state_dir: Optional[Path] = None):
        """Initialize the delta state manager.

        Args:
            state_dir: Directory to store delta state files.
                      Defaults to ~/.teams-exporter/delta/
        """
        if state_dir is None:
            state_dir = Path.home() / ".teams-exporter" / "delta"
        self.state_dir = Path(state_dir)
        self.state_dir.mkdir(parents=True, exist_ok=True)

    def _get_state_file(self, chat_id: str) -> Path:
        """Get the state file path for a chat.

        Args:
            chat_id: The chat ID

        Returns:
            Path to the state file
        """
        # Sanitize chat ID for filesystem
        safe_id = chat_id.replace("/", "_").replace(":", "_")
        return self.state_dir / f"{safe_id}.json"

    def get_state(self, chat_id: str) -> Optional[Dict[str, Any]]:
        """Get the sync state for a chat.

        Args:
            chat_id: The chat ID

        Returns:
            State dict with delta_link, last_message_id, last_sync_time, etc.
            Returns None if no state exists.
        """
        state_file = self._get_state_file(chat_id)
        if not state_file.exists():
            return None

        try:
            with open(state_file, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            # Corrupted state file, remove it
            state_file.unlink(missing_ok=True)
            return None

    def save_state(
        self,
        chat_id: str,
        delta_link: Optional[str] = None,
        last_message_id: Optional[str] = None,
        last_message_time: Optional[str] = None,
        message_count: int = 0
    ) -> None:
        """Save the sync state for a chat.

        Args:
            chat_id: The chat ID
            delta_link: The delta link for the next sync (from Graph API)
            last_message_id: ID of the last processed message
            last_message_time: Timestamp of the last processed message
            message_count: Total number of messages processed
        """
        state_file = self._get_state_file(chat_id)

        state = {
            "chat_id": chat_id,
            "delta_link": delta_link,
            "last_message_id": last_message_id,
            "last_message_time": last_message_time,
            "message_count": message_count,
            "last_sync": datetime.now(timezone.utc).isoformat()
        }

        with open(state_file, "w") as f:
            json.dump(state, f, indent=2)

    def clear_state(self, chat_id: str) -> None:
        """Clear the sync state for a chat.

        Args:
            chat_id: The chat ID
        """
        state_file = self._get_state_file(chat_id)
        state_file.unlink(missing_ok=True)

    def clear_all_states(self) -> None:
        """Clear all sync states."""
        if self.state_dir.exists():
            for state_file in self.state_dir.glob("*.json"):
                state_file.unlink()

    def list_states(self) -> Dict[str, Dict[str, Any]]:
        """List all saved states.

        Returns:
            Dict mapping chat IDs to their states
        """
        states = {}
        if self.state_dir.exists():
            for state_file in self.state_dir.glob("*.json"):
                try:
                    with open(state_file, "r") as f:
                        state = json.load(f)
                        if "chat_id" in state:
                            states[state["chat_id"]] = state
                except (json.JSONDecodeError, IOError):
                    # Skip corrupted files
                    continue
        return states

    def has_state(self, chat_id: str) -> bool:
        """Check if a chat has saved state.

        Args:
            chat_id: The chat ID

        Returns:
            True if state exists, False otherwise
        """
        return self._get_state_file(chat_id).exists()