"""Working out which OneDrive or SharePoint library a tool call means.

A `drive` argument can be:
  * "onedrive" (or empty): the signed-in person's own OneDrive
  * a SharePoint site URL, e.g. https://contoso.sharepoint.com/sites/Finance,
    meaning that site's default document library
  * a drive id, as returned by the list_libraries tool
"""

from __future__ import annotations

from urllib.parse import quote, urlparse

from .graph import Graph, GraphError

_site_cache: dict[str, str] = {}


async def site_id_from_url(graph: Graph, url: str) -> str:
    if url in _site_cache:
        return _site_cache[url]
    parsed = urlparse(url)
    parts = [p for p in parsed.path.split("/") if p]
    if len(parts) >= 2 and parts[0] in ("sites", "teams"):
        site_path = f"/{parts[0]}/{parts[1]}"
        site = await graph.get(f"/sites/{parsed.hostname}:{quote(site_path)}")
    else:
        site = await graph.get(f"/sites/{parsed.hostname}")
    _site_cache[url] = site["id"]
    return site["id"]


async def drive_base(graph: Graph, drive: str | None) -> str:
    d = (drive or "").strip()
    if d.lower() in ("", "onedrive", "me", "my onedrive"):
        return "/me/drive"
    if d.startswith("https://"):
        return f"/sites/{await site_id_from_url(graph, d)}/drive"
    return f"/drives/{quote(d, safe='!')}"


def clean_path(path: str | None) -> str:
    return "/".join(p for p in (path or "").replace("\\", "/").split("/") if p)


def item_ref(base: str, path: str | None) -> str:
    """The Graph URL of an item by its path inside a drive."""
    p = clean_path(path)
    return f"{base}/root" if not p else f"{base}/root:/{quote(p)}:"


def children_ref(base: str, path: str | None) -> str:
    p = clean_path(path)
    return f"{base}/root/children" if not p else f"{base}/root:/{quote(p)}:/children"


async def get_item(graph: Graph, drive: str | None, path: str) -> tuple[str, dict]:
    base = await drive_base(graph, drive)
    try:
        item = await graph.get(item_ref(base, path))
    except GraphError as exc:
        if exc.status == 404:
            raise GraphError(404, f"Nothing found at '{path}' in {drive or 'OneDrive'}.", "itemNotFound") from None
        raise
    # Always address the item by its own drive and id from here on.
    return f"/drives/{item['parentReference']['driveId']}/items/{item['id']}", item


def describe(item: dict) -> dict:
    out = {
        "name": item.get("name"),
        "type": "folder" if "folder" in item else "file",
        "id": item.get("id"),
        "size": item.get("size"),
        "modified": item.get("lastModifiedDateTime"),
        "web_url": item.get("webUrl"),
    }
    parent = item.get("parentReference") or {}
    if parent.get("path"):
        out["folder"] = parent["path"].split("root:", 1)[-1] or "/"
    if parent.get("driveId"):
        out["drive_id"] = parent["driveId"]
    if "folder" in item:
        out["child_count"] = item["folder"].get("childCount")
    return out
