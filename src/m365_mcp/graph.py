"""A small async Microsoft Graph client: retries, paging and uploads."""

from __future__ import annotations

import asyncio
from typing import Any, Callable

import httpx

GRAPH = "https://graph.microsoft.com/v1.0"
RETRY_STATUSES = {429, 502, 503, 504}
MAX_ATTEMPTS = 4
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024
# Upload-session chunks must be a multiple of 320 KiB.
CHUNK_SIZE = 10 * 320 * 1024


class GraphError(Exception):
    def __init__(self, status: int, message: str, code: str = ""):
        super().__init__(f"Microsoft Graph error {status}{f' ({code})' if code else ''}: {message}")
        self.status = status
        self.code = code


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
            if resp.status_code in RETRY_STATUSES and attempt < MAX_ATTEMPTS:
                delay = float(resp.headers.get("Retry-After", 2 ** attempt))
                await asyncio.sleep(min(delay, 30))
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
