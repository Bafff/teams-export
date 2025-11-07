from __future__ import annotations

import time
from typing import Callable, Dict, Iterable, Iterator, List, Optional

import requests

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
DEFAULT_TIMEOUT = 60
MAX_RETRIES = 4
INITIAL_RETRY_DELAY = 2.0  # seconds


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

    def _request_with_retry(
        self,
        url: str,
        params: Dict[str, str] | None = None,
    ) -> requests.Response:
        """Execute a GET request with exponential backoff retry on rate limiting."""
        last_exception = None

        for attempt in range(MAX_RETRIES):
            try:
                resp = self._session.get(url, params=params, timeout=DEFAULT_TIMEOUT)

                # Handle rate limiting (429) with retry
                if resp.status_code == 429:
                    retry_after = resp.headers.get("Retry-After")
                    if retry_after:
                        try:
                            wait_time = int(retry_after)
                        except ValueError:
                            wait_time = INITIAL_RETRY_DELAY * (2 ** attempt)
                    else:
                        wait_time = INITIAL_RETRY_DELAY * (2 ** attempt)

                    if attempt < MAX_RETRIES - 1:
                        print(f"Rate limited. Waiting {wait_time}s before retry {attempt + 1}/{MAX_RETRIES}...")
                        time.sleep(wait_time)
                        continue
                    else:
                        raise GraphError(self._format_error(resp))

                # Handle other 5xx errors with retry
                if 500 <= resp.status_code < 600:
                    if attempt < MAX_RETRIES - 1:
                        wait_time = INITIAL_RETRY_DELAY * (2 ** attempt)
                        print(f"Server error {resp.status_code}. Retrying in {wait_time}s...")
                        time.sleep(wait_time)
                        continue

                # Success or non-retryable error
                return resp

            except requests.exceptions.RequestException as exc:
                last_exception = exc
                if attempt < MAX_RETRIES - 1:
                    wait_time = INITIAL_RETRY_DELAY * (2 ** attempt)
                    print(f"Network error: {exc}. Retrying in {wait_time}s...")
                    time.sleep(wait_time)
                    continue

        # If we exhausted retries
        if last_exception:
            raise GraphError(f"Request failed after {MAX_RETRIES} attempts: {last_exception}")
        raise GraphError(f"Request failed after {MAX_RETRIES} attempts")

    def _paginate(
        self,
        url: str,
        params: Dict[str, str] | None = None,
        *,
        stop_condition: Optional[Callable[[dict], bool]] = None,
        progress_callback: Optional[Callable[[int], None]] = None,
        max_items: Optional[int] = None,
    ) -> Iterator[dict]:
        """Paginate through API results with optional progress tracking and limits.

        Args:
            url: API endpoint URL
            params: Query parameters for first request
            stop_condition: Function that returns True to stop iteration
            progress_callback: Called with count after each page is fetched
            max_items: Maximum number of items to fetch (None = unlimited)
        """
        count = 0
        while url:
            resp = self._request_with_retry(url, params=params)
            params = None  # Only include params on first request.
            if resp.status_code >= 400:
                raise GraphError(self._format_error(resp))
            payload = resp.json()
            for item in payload.get("value", []):
                yield item
                count += 1
                if stop_condition and stop_condition(item):
                    return
                if max_items and count >= max_items:
                    return

            if progress_callback:
                progress_callback(count)

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

    def list_chats(
        self,
        *,
        limit: Optional[int] = None,
        progress_callback: Optional[Callable[[int], None]] = None,
    ) -> List[dict]:
        """List accessible chats with optional limit and progress tracking.

        Args:
            limit: Maximum number of chats to fetch (None = all chats)
            progress_callback: Function called with count after each page

        Returns:
            List of chat objects with expanded members
        """
        url = f"{self._base_url}/me/chats"
        params = {
            "$expand": "members",
            "$top": "50",  # Fetch 50 chats per request
        }
        return list(self._paginate(
            url,
            params=params,
            max_items=limit,
            progress_callback=progress_callback,
        ))

    def list_chat_messages(
        self,
        chat_id: str,
        *,
        stop_condition: Optional[Callable[[dict], bool]] = None,
    ) -> List[dict]:
        url = f"{self._base_url}/me/chats/{chat_id}/messages"
        params = {
            "$top": "100",  # Increased from 50 for better performance
        }
        return list(self._paginate(url, params=params, stop_condition=stop_condition))

    def close(self) -> None:
        self._session.close()

    def __enter__(self) -> "GraphClient":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # pragma: no cover - cleanup path
        self.close()
