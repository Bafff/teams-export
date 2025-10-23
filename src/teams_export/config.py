from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

DEFAULT_CONFIG_DIR = Path("~/.teams-exporter").expanduser()
DEFAULT_CONFIG_PATH = DEFAULT_CONFIG_DIR / "config.json"
DEFAULT_TOKEN_CACHE_PATH = DEFAULT_CONFIG_DIR / "token_cache.json"
DEFAULT_AUTHORITY = "https://login.microsoftonline.com/common"
DEFAULT_SCOPES = ["Chat.Read", "Chat.ReadBasic", "Chat.ReadWrite"]


class ConfigError(RuntimeError):
    """Raised when mandatory configuration is missing."""


@dataclass(slots=True)
class AppConfig:
    client_id: str
    authority: str = DEFAULT_AUTHORITY
    scopes: list[str] = None
    token_cache_path: Path = DEFAULT_TOKEN_CACHE_PATH

    def __post_init__(self) -> None:
        if not self.scopes:
            self.scopes = list(DEFAULT_SCOPES)
        # Normalise to expanded absolute path.
        self.token_cache_path = self.token_cache_path.expanduser()


def _load_file_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def load_config(path: Path | None = None) -> AppConfig:
    """Load CLI configuration, falling back to defaults and env overrides."""

    cfg_path = path or DEFAULT_CONFIG_PATH
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

    token_cache_path = Path(raw.get("token_cache_path", DEFAULT_TOKEN_CACHE_PATH))

    return AppConfig(
        client_id=client_id,
        authority=authority,
        scopes=scopes,
        token_cache_path=token_cache_path,
    )


def ensure_config_dir() -> Path:
    DEFAULT_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    return DEFAULT_CONFIG_DIR
