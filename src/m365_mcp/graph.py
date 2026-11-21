"""A small async Microsoft Graph client: retries, paging and uploads."""

from __future__ import annotations

import asyncio
import email.utils
import re
import time
from typing import Any, Callable

import httpx

GRAPH = "https://graph.microsoft.com/v1.0"
RETRY_STATUSES = {429, 502, 503, 504}
# Only these methods are safe to repeat after a gateway error: the first attempt may already
# have worked (a POST to /send that times out can still send the email), so other methods
# are only retried when Graph says it didn't run them (429, or 503 with Retry-After).
IDEMPOTENT_METHODS = {"GET", "HEAD", "PUT", "DELETE"}
# Outlook, calendar and drive ids are URL-safe base64-like strings. Checking them stops an id
# like "<draftId>/send?" from turning one request into a different one.
_ID_RE = re.compile(r"^[A-Za-z0-9_=+!-]{1,1024}$")
MAX_ATTEMPTS = 4
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024
# Upload-session chunks must be a multiple of 320 KiB.
CHUNK_SIZE = 10 * 320 * 1024


class GraphError(Exception):
    def __init__(self, status: int, message: str, code: str = ""):
        super().__init__(f"Microsoft Graph error {status}{f' ({code})' if code else ''}: {message}")
        self.status = status
        self.code = code


def safe_id(value: str, what: str = "id") -> str:
    """`value` if it's a well-formed Graph id, ready to put in a URL; otherwise a clear error."""
    v = (value or "").strip()
    if not _ID_RE.match(v):
        raise ValueError(f"'{value}' isn't a valid {what}. Use the id exactly as a search or list tool returned it.")
    return v


def _retry_delay(resp: httpx.Response, attempt: int) -> float:
    value = resp.headers.get("Retry-After")
    if value:
        try:
            return float(value)
        except ValueError:
            try:
                return max(0.0, email.utils.parsedate_to_datetime(value).timestamp() - time.time())
            except (TypeError, ValueError):
                pass
    return float(2 ** attempt)


def _should_retry(method: str, resp: httpx.Response) -> bool:
    if resp.status_code == 429:
        return True
    if resp.status_code not in RETRY_STATUSES:
        return False
    if method.upper() in IDEMPOTENT_METHODS:
        return True
    return resp.status_code == 503 and "Retry-After" in resp.headers


class Graph:
    def __init__(self, token_provider: Callable[[], str], client: httpx.AsyncClient | None = None):
        self._token = token_provider
        self._client = client or httpx.AsyncClient(timeout=httpx.Timeout(120.0), follow_redirects=True)

    async def aclose(self) -> None:
        await self._client.aclose()

    def _url(self, path: str) -> str:
        return path if path.startswith("https://") else GRAPH + path

    async def request(
        self,
        method: str,
        path: str,
        *,
        params: dict | None = None,
        json: Any = None,
        content: bytes | None = None,
        headers: dict | None = None,
        auth: bool = True,
    ) -> httpx.Response:
        hdrs = dict(headers or {})
        if auth:
            hdrs["Authorization"] = f"Bearer {self._token()}"
        for attempt in range(1, MAX_ATTEMPTS + 1):
            resp = await self._client.request(
                method, self._url(path), params=params, json=json, content=content, headers=hdrs
            )
            if attempt < MAX_ATTEMPTS and _should_retry(method, resp):
                await asyncio.sleep(min(_retry_delay(resp, attempt), 30))
                continue
            break
        if resp.status_code >= 400:
            code, message = "", resp.text[:500]
            try:
                err = resp.json().get("error", {})
                code, message = err.get("code", ""), err.get("message", message)
            except ValueError:
                pass
            raise GraphError(resp.status_code, message, code)
        return resp

    async def get(self, path: str, **kw) -> Any:
        return (await self.request("GET", path, **kw)).json()

    async def post(self, path: str, json: Any = None, **kw) -> Any:
        resp = await self.request("POST", path, json=json, **kw)
        return resp.json() if resp.content else {}

    async def patch(self, path: str, json: Any, **kw) -> Any:
        resp = await self.request("PATCH", path, json=json, **kw)
        return resp.json() if resp.content else {}

    async def delete(self, path: str) -> None:
        await self.request("DELETE", path)

    async def get_bytes(self, path: str, params: dict | None = None) -> bytes:
        return (await self.request("GET", path, params=params)).content

    async def get_all(self, path: str, params: dict | None = None, limit: int = 500, **kw) -> list:
        items: list = []
        page = await self.get(path, params=params, **kw)
        while True:
            items.extend(page.get("value", []))
            nxt = page.get("@odata.nextLink")
            if not nxt or len(items) >= limit:
                return items[:limit]
            page = await self.get(nxt, **kw)

    async def upload_to_session(self, upload_url: str, data: bytes) -> dict:
        """PUT data to a pre-authorised upload session URL in chunks."""
        total = len(data)
        result: dict = {}
        for start in range(0, total, CHUNK_SIZE):
            chunk = data[start : start + CHUNK_SIZE]
            end = start + len(chunk) - 1
            resp = await self.request(
                "PUT",
                upload_url,
                content=chunk,
                headers={"Content-Length": str(len(chunk)), "Content-Range": f"bytes {start}-{end}/{total}"},
                auth=False,
            )
            if resp.content:
                result = resp.json()
        return result
