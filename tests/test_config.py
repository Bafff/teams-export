"""Tests for configuration loading."""

import json
from pathlib import Path

import pytest

from teams_export.config import load_config, ConfigError


class TestConfigLoading:
    """Test configuration loading from various sources."""

    def test_load_from_file(self, tmp_path, monkeypatch):
        """Test loading config from JSON file."""
        config_dir = tmp_path / ".teams-exporter"
        config_dir.mkdir()
        config_file = config_dir / "config.json"

        config_data = {
            "client_id": "test-client-123",
            "authority": "https://login.microsoftonline.com/tenant",
            "scopes": ["Chat.Read", "User.Read"],
            "token_cache_path": str(config_dir / "cache.json")
        }
        config_file.write_text(json.dumps(config_data))

        # Patch home directory to use tmp_path
        monkeypatch.setattr(Path, "home", lambda: tmp_path)

        config = load_config()
        assert config.client_id == "test-client-123"
        assert config.authority == "https://login.microsoftonline.com/tenant"
        assert "Chat.Read" in config.scopes
        assert "User.Read" in config.scopes

    def test_load_from_env_vars(self, monkeypatch):
        """Test loading config from environment variables."""
        monkeypatch.setenv("TEAMS_EXPORT_CLIENT_ID", "env-client-456")
        monkeypatch.setenv("TEAMS_EXPORT_AUTHORITY", "https://login.microsoftonline.com/env-tenant")
        monkeypatch.setenv("TEAMS_EXPORT_SCOPES", "Chat.Read,User.Read,Chat.ReadWrite")

        config = load_config()
        assert config.client_id == "env-client-456"
        assert config.authority == "https://login.microsoftonline.com/env-tenant"
        assert config.scopes == ["Chat.Read", "User.Read", "Chat.ReadWrite"]

    def test_env_vars_override_file(self, tmp_path, monkeypatch):
        """Test that environment variables override file config."""
        config_dir = tmp_path / ".teams-exporter"
        config_dir.mkdir()
        config_file = config_dir / "config.json"

        config_data = {
            "client_id": "file-client",
            "authority": "https://login.microsoftonline.com/file-tenant"
        }
        config_file.write_text(json.dumps(config_data))

        monkeypatch.setattr(Path, "home", lambda: tmp_path)
        monkeypatch.setenv("TEAMS_EXPORT_CLIENT_ID", "env-client")

        config = load_config()
        assert config.client_id == "env-client"  # From env
        assert config.authority == "https://login.microsoftonline.com/file-tenant"  # From file

    def test_default_values(self, tmp_path, monkeypatch):
        """Test default values when no config is provided."""
        # Create empty config file
        config_dir = tmp_path / ".teams-exporter"
        config_dir.mkdir()
        config_file = config_dir / "config.json"
        config_file.write_text("{}")

        monkeypatch.setattr(Path, "home", lambda: tmp_path)
        monkeypatch.setenv("TEAMS_EXPORT_CLIENT_ID", "default-client")

        config = load_config()
        # Should have default authority and scopes
        assert config.authority == "https://login.microsoftonline.com/common"
        assert "Chat.Read" in config.scopes
        assert "Chat.ReadBasic" in config.scopes

    def test_missing_client_id_raises_error(self, tmp_path, monkeypatch):
        """Test that missing client_id raises ConfigError."""
        # Create config without client_id
        config_dir = tmp_path / ".teams-exporter"
        config_dir.mkdir()
        config_file = config_dir / "config.json"
        config_file.write_text('{"authority": "test"}')

        monkeypatch.setattr(Path, "home", lambda: tmp_path)

        with pytest.raises(ConfigError, match="client_id"):
            load_config()

    def test_invalid_json_raises_error(self, tmp_path, monkeypatch):
        """Test that invalid JSON raises ConfigError."""
        config_dir = tmp_path / ".teams-exporter"
        config_dir.mkdir()
        config_file = config_dir / "config.json"
        config_file.write_text("not valid json {")

        monkeypatch.setattr(Path, "home", lambda: tmp_path)

        with pytest.raises(ConfigError, match="Invalid JSON"):
            load_config()

    def test_token_cache_path_expansion(self, tmp_path, monkeypatch):
        """Test that ~ in token_cache_path is expanded."""
        config_dir = tmp_path / ".teams-exporter"
        config_dir.mkdir()
        config_file = config_dir / "config.json"

        config_data = {
            "client_id": "test-client",
            "token_cache_path": "~/.teams-exporter/custom_cache.json"
        }
        config_file.write_text(json.dumps(config_data))

        monkeypatch.setattr(Path, "home", lambda: tmp_path)

        config = load_config()
        assert config.token_cache_path == (tmp_path / ".teams-exporter" / "custom_cache.json")
