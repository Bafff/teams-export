from __future__ import annotations

from typing import Dict, Iterable, Iterator, List

import requests

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
DEFAULT_TIMEOUT = 30


class GraphError(RuntimeError):
    """Raised when the Graph API returns an error."""


class GraphClient:
    def __init__(self, token: str, base_url: str = GRAPH_BASE_URL) -> None:
        self._session = requests.Session()
        self._session.headers.update(
            {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
                "Content-Type": "application/json",
            }
        )
        self._base_url = base_url.rstrip("/")

    def _paginate(self, url: str, params: Dict[str, str] | None = None) -> Iterator[dict]:
        while url:
            resp = self._session.get(url, params=params, timeout=DEFAULT_TIMEOUT)
            params = None  # Only include params on first request.
            if resp.status_code >= 400:
                raise GraphError(self._format_error(resp))
            payload = resp.json()
            for item in payload.get("value", []):
                yield item
            url = payload.get("@odata.nextLink")

    def _format_error(self, response: requests.Response) -> str:
        try:
            detail = response.json()
        except ValueError:
            detail = {"error": response.text}
        base = detail.get("error") if isinstance(detail, dict) else detail
        if isinstance(base, dict):
            message = base.get("message")
            code = base.get("code")
            return f"Graph API error {code or response.status_code}: {message}"
        return f"Graph API error {response.status_code}: {base}"

    def list_chats(self) -> List[dict]:
        url = f"{self._base_url}/me/chats"
        params = {
            "$expand": "members($select=displayName,email,userId)",
        }
        return list(self._paginate(url, params=params))

    def list_chat_messages(
        self,
        chat_id: str,
        start_iso: str,
        end_iso: str,
    ) -> List[dict]:
        url = f"{self._base_url}/me/chats/{chat_id}/messages"
        params = {
            "$filter": f"lastModifiedDateTime ge {start_iso} and lastModifiedDateTime le {end_iso}",
            "$orderby": "lastModifiedDateTime asc",
        }
        return list(self._paginate(url, params=params))

    def close(self) -> None:
        self._session.close()

    def __enter__(self) -> "GraphClient":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # pragma: no cover - cleanup path
        self.close()
