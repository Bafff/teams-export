from __future__ import annotations

from pathlib import Path
from typing import Callable, Optional

import msal

from .config import AppConfig, ensure_config_dir


class AuthError(RuntimeError):
    """Raised when authentication fails."""


def _load_cache(path: Path) -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if path.exists():
        cache.deserialize(path.read_text(encoding="utf-8"))
    return cache


def _save_cache(cache: msal.SerializableTokenCache, path: Path) -> None:
    if cache.has_state_changed:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(cache.serialize(), encoding="utf-8")


def acquire_token(
    config: AppConfig,
    *,
    message_callback: Optional[Callable[[str], None]] = None,
    force_refresh: bool = False,
) -> str:
    """Authenticate the user and return an access token."""

    ensure_config_dir()
    cache = _load_cache(config.token_cache_path)
    app = msal.PublicClientApplication(
        client_id=config.client_id,
        authority=config.authority,
        token_cache=cache,
    )

    if not force_refresh:
        accounts = app.get_accounts()
        if accounts:
            for account in accounts:
                result = app.acquire_token_silent(config.scopes, account=account)
                if result and "access_token" in result:
                    _save_cache(cache, config.token_cache_path)
                    return result["access_token"]

    flow = app.initiate_device_flow(scopes=config.scopes)
    if "user_code" not in flow:
        raise AuthError("Device flow failed to initialise; check client_id and scopes.")

    if message_callback:
        message_callback(flow["message"])  # MSAL supplies a human-friendly instruction string.

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise AuthError(result.get("error_description") or "Authentication failed.")

    _save_cache(cache, config.token_cache_path)
    return result["access_token"]
