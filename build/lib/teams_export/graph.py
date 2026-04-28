from __future__ import annotations

import time
from typing import Callable, Dict, Iterable, Iterator, List, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

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
            List of chat objects with expanded members and lastMessagePreview
        """
        url = f"{self._base_url}/me/chats"
        params = {
            "$expand": "members,lastMessagePreview",
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
            "$top": "50",  # Graph API maximum for chat messages endpoint
        }
        return list(self._paginate(url, params=params, stop_condition=stop_condition))

    def list_chat_messages_delta(
        self,
        chat_id: str,
        delta_link: Optional[str] = None,
        *,
        progress_callback: Optional[Callable[[int], None]] = None,
    ) -> Tuple[List[dict], Optional[str]]:
        """List chat messages using delta query for incremental sync.

        Args:
            chat_id: The chat ID
            delta_link: Delta link from previous sync (if any)
            progress_callback: Function called with count after each page

        Returns:
            Tuple of (messages list, next delta link)
        """
        if delta_link:
            # Use the provided delta link directly
            url = delta_link
            params = None
        else:
            # Initial delta query
            url = f"{self._base_url}/me/chats/{chat_id}/messages/delta"
            params = {
                "$top": "50",  # Graph API maximum for chat messages endpoint
            }

        messages = []
        next_delta_link = None
        count = 0

        while url:
            resp = self._request_with_retry(url, params=params)
            params = None  # Only include params on first request
            if resp.status_code >= 400:
                raise GraphError(self._format_error(resp))

            payload = resp.json()

            # Collect messages
            for item in payload.get("value", []):
                messages.append(item)
                count += 1

            if progress_callback:
                progress_callback(count)

            # Check for next page or delta link
            url = payload.get("@odata.nextLink")
            if not url:
                # No more pages, check for delta link
                next_delta_link = payload.get("@odata.deltaLink")

        return messages, next_delta_link

    def list_chat_messages_parallel(
        self,
        chat_id: str,
        *,
        stop_condition: Optional[Callable[[dict], bool]] = None,
        max_workers: int = 3,
        progress_callback: Optional[Callable[[int], None]] = None,
    ) -> List[dict]:
        """List chat messages with parallel pagination for faster fetching.

        This method fetches the first page to understand the total structure,
        then fetches remaining pages in parallel.

        Args:
            chat_id: The chat ID
            stop_condition: Function that returns True to stop iteration
            max_workers: Maximum number of parallel workers (default 3)
            progress_callback: Function called with count after each page

        Returns:
            List of message dictionaries, sorted by timestamp
        """
        url = f"{self._base_url}/me/chats/{chat_id}/messages"
        params = {
            "$top": "50",  # Graph API maximum for chat messages endpoint
        }

        # Fetch first page to get initial data and next link
        resp = self._request_with_retry(url, params=params)
        if resp.status_code >= 400:
            raise GraphError(self._format_error(resp))

        payload = resp.json()
        all_messages = []
        page_urls = []

        # Process first page
        for item in payload.get("value", []):
            all_messages.append(item)
            if stop_condition and stop_condition(item):
                return all_messages

        # Collect all page URLs
        next_url = payload.get("@odata.nextLink")
        while next_url:
            page_urls.append(next_url)
            # Estimate number of pages (can't know exactly without fetching)
            # Break after collecting reasonable number of URLs for parallel fetch
            if len(page_urls) >= 20:  # Reasonable limit for initial batch
                break
            next_url = None  # We'll discover more URLs as we fetch

        if not page_urls:
            # No more pages, return what we have
            return all_messages

        # Thread-safe counter for progress
        message_count = len(all_messages)
        count_lock = threading.Lock()

        def fetch_page(page_url):
            """Fetch a single page of messages."""
            resp = self._request_with_retry(page_url)
            if resp.status_code >= 400:
                raise GraphError(self._format_error(resp))

            page_data = resp.json()
            messages = []
            for item in page_data.get("value", []):
                messages.append(item)
                if stop_condition and stop_condition(item):
                    return messages, None  # Signal to stop

            next_link = page_data.get("@odata.nextLink")
            return messages, next_link

        # Use ThreadPoolExecutor for parallel fetching
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit initial batch of URLs
            future_to_url = {executor.submit(fetch_page, url): url for url in page_urls}

            while future_to_url:
                # Process completed futures
                done_futures = []
                for future in as_completed(future_to_url):
                    done_futures.append(future)
                    url = future_to_url[future]

                    try:
                        messages, next_link = future.result()
                        all_messages.extend(messages)

                        with count_lock:
                            message_count += len(messages)
                            if progress_callback:
                                progress_callback(message_count)

                        # If we got a next link, add it to the queue
                        if next_link and next_link not in future_to_url.values():
                            new_future = executor.submit(fetch_page, next_link)
                            future_to_url[new_future] = next_link

                        # Check if we should stop
                        if next_link is None and stop_condition:
                            # One of the pages triggered stop condition
                            # Cancel remaining futures
                            for f in future_to_url:
                                f.cancel()
                            break

                    except Exception as e:
                        print(f"Error fetching page: {e}")

                    # Remove processed future
                    del future_to_url[future]

                # Break if all futures are done
                if not future_to_url:
                    break

        # Sort messages by timestamp (Graph API returns newest first, we want consistent order)
        all_messages.sort(
            key=lambda m: m.get("createdDateTime") or m.get("lastModifiedDateTime") or "",
            reverse=True  # Keep newest first to match regular pagination
        )

        return all_messages

    def close(self) -> None:
        self._session.close()

    def __enter__(self) -> "GraphClient":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # pragma: no cover - cleanup path
        self.close()
