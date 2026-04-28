from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

DEFAULT_AUTHORITY = "https://login.microsoftonline.com/common"
DEFAULT_SCOPES = ["Chat.Read", "Chat.ReadBasic", "Chat.ReadWrite"]


def _default_config_dir() -> Path:
    return Path.home() / ".teams-exporter"


def _default_config_path() -> Path:
    return _default_config_dir() / "config.json"


def _default_token_cache_path() -> Path:
    return _default_config_dir() / "token_cache.json"


class ConfigError(RuntimeError):
    """Raised when mandatory configuration is missing."""


@dataclass(slots=True)
class AppConfig:
    client_id: str
    authority: str = DEFAULT_AUTHORITY
    scopes: list[str] | None = None
    token_cache_path: Path | None = None

    def __post_init__(self) -> None:
        if not self.scopes:
            self.scopes = list(DEFAULT_SCOPES)
        if self.token_cache_path is None:
            self.token_cache_path = _default_token_cache_path()
        else:
            resolved = Path(self.token_cache_path)
            text = str(resolved)
            if text.startswith("~"):
                resolved = Path(str(Path.home()) + text[1:])
            self.token_cache_path = resolved.expanduser()


def _load_file_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as handle:
            return json.load(handle)
    except json.JSONDecodeError as exc:
        raise ConfigError(f"Invalid JSON in {path}: {exc}") from exc


def load_config(path: Path | None = None) -> AppConfig:
    """Load CLI configuration, falling back to defaults and env overrides."""

    cfg_path = path or _default_config_path()
    raw = _load_file_config(cfg_path)

    env_client_id = os.environ.get("TEAMS_EXPORT_CLIENT_ID")
    env_authority = os.environ.get("TEAMS_EXPORT_AUTHORITY")
    env_scopes = os.environ.get("TEAMS_EXPORT_SCOPES")

    client_id = env_client_id or raw.get("client_id")
    if not client_id:
        raise ConfigError(
            "Missing client_id; set TEAMS_EXPORT_CLIENT_ID or define it in"
            f" {cfg_path}."
        )

    authority = env_authority or raw.get("authority", DEFAULT_AUTHORITY)

    scopes: list[str] | None = None
    if env_scopes:
        scopes = [scope.strip() for scope in env_scopes.split(",") if scope.strip()]
    else:
        scopes = raw.get("scopes")

    token_cache_value = raw.get("token_cache_path")
    token_cache_path = (
        Path(token_cache_value)
        if token_cache_value
        else _default_token_cache_path()
    )

    return AppConfig(
        client_id=client_id,
        authority=authority,
        scopes=scopes,
        token_cache_path=token_cache_path,
    )


def ensure_config_dir() -> Path:
    config_dir = _default_config_dir()
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir
