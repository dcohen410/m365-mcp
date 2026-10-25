"""OneDrive and SharePoint file operations."""

from __future__ import annotations

import io
from pathlib import PurePosixPath
from urllib.parse import quote

from pypdf import PdfReader

from . import drives, localfs
from .config import Config
from .graph import SIMPLE_UPLOAD_LIMIT, Graph, GraphError

TEXT_EXTENSIONS = {
    ".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".log", ".yaml", ".yml", ".ini", ".py", ".js",
}
# Formats Microsoft Graph can convert to PDF, which we then read as text.
CONVERTIBLE_TO_PDF = {
    ".doc", ".docx", ".rtf", ".odt", ".ppt", ".pptx", ".pps", ".ppsx", ".odp", ".eml", ".msg",
}


def pdf_text(data: bytes, max_chars: int) -> str:
    reader = PdfReader(io.BytesIO(data))
    parts, size = [], 0
    for n, page in enumerate(reader.pages, 1):
        text = page.extract_text() or ""
        parts.append(f"--- page {n} ---\n{text}")
        size += len(text)
        if size >= max_chars:
            break
    return "\n".join(parts)[:max_chars]


async def search_sites(graph: Graph, query: str) -> list[dict]:
    sites = await graph.get_all("/sites", params={"search": query}, limit=50)
    return [{"name": s.get("displayName"), "url": s.get("webUrl"), "id": s.get("id")} for s in sites]


async def list_libraries(graph: Graph, site_url: str | None) -> list[dict]:
    if site_url:
        site_id = await drives.site_id_from_url(graph, site_url)
        found = await graph.get_all(f"/sites/{site_id}/drives")
    else:
        found = [await graph.get("/me/drive")]
    return [
        {"name": d.get("name"), "drive_id": d.get("id"), "web_url": d.get("webUrl"), "type": d.get("driveType")}
        for d in found
    ]


async def list_folder(graph: Graph, drive: str | None, path: str, limit: int = 200) -> list[dict]:
    base = await drives.drive_base(graph, drive)
    items = await graph.get_all(drives.children_ref(base, path), limit=limit)
    return [drives.describe(i) for i in items]


async def search_files(graph: Graph, query: str, drive: str | None, limit: int = 25) -> list[dict]:
    if drive:
        base = await drives.drive_base(graph, drive)
        q = query.replace("'", "''")
        items = await graph.get_all(f"{base}/root/search(q='{quote(q)}')", limit=limit)
        return [drives.describe(i) for i in items]
    body = {"requests": [{"entityTypes": ["driveItem"], "query": {"queryString": query}, "size": limit}]}
    result = await graph.post("/search/query", json=body)
    hits = []
    for container in result.get("value", []):
        for hc in container.get("hitsContainers", []):
            for hit in hc.get("hits", []):
                hits.append(drives.describe(hit.get("resource", {})))
    return hits[:limit]


async def read_file(graph: Graph, drive: str | None, path: str, max_chars: int = 50_000) -> str:
    ref, item = await drives.get_item(graph, drive, path)
    if "folder" in item:
        raise ValueError(f"'{path}' is a folder; use list_folder instead.")
    ext = PurePosixPath(item["name"]).suffix.lower()
    if ext in TEXT_EXTENSIONS:
        data = await graph.get_bytes(f"{ref}/content")
        return data.decode("utf-8", errors="replace")[:max_chars]
    if ext == ".pdf":
        return pdf_text(await graph.get_bytes(f"{ref}/content"), max_chars)
    if ext in CONVERTIBLE_TO_PDF:
        return pdf_text(await graph.get_bytes(f"{ref}/content", params={"format": "pdf"}), max_chars)
    if ext in (".xlsx", ".xlsm", ".xls"):
        raise ValueError("This is a spreadsheet; use excel_list_sheets and excel_read instead.")
    raise ValueError(f"Can't read {ext or 'this'} files as text. Use download_file to save a copy instead.")


async def download_file(graph: Graph, cfg: Config, drive: str | None, path: str, local_folder: str) -> str:
    ref, item = await drives.get_item(graph, drive, path)
    if "folder" in item:
        raise ValueError(f"'{path}' is a folder; download files one at a time.")
    data = await graph.get_bytes(f"{ref}/content")
    saved = localfs.write_file(local_folder, item["name"], data, cfg.local_roots())
    return str(saved)


async def upload_bytes(
    graph: Graph, drive: str | None, folder: str, filename: str, data: bytes, overwrite: bool = False
) -> dict:
    """Upload into folder. Without overwrite, a clash gets a new name like 'name 1.pdf'."""
    base = await drives.drive_base(graph, drive)
    behaviour = "replace" if overwrite else "rename"
    target = drives.clean_path(f"{drives.clean_path(folder)}/{localfs.safe_filename(filename)}")
    if len(data) <= SIMPLE_UPLOAD_LIMIT:
        resp = await graph.request(
            "PUT",
            f"{base}/root:/{quote(target)}:/content",
            params={"@microsoft.graph.conflictBehavior": behaviour},
            content=data,
            headers={"Content-Type": "application/octet-stream"},
        )
        return drives.describe(resp.json())
    session = await graph.post(
        f"{base}/root:/{quote(target)}:/createUploadSession",
        json={"item": {"@microsoft.graph.conflictBehavior": behaviour}},
    )
    return drives.describe(await graph.upload_to_session(session["uploadUrl"], data))


async def upload_file(
    graph: Graph, cfg: Config, local_path: str, drive: str | None, folder: str,
    new_name: str | None = None, overwrite: bool = False,
) -> dict:
    name, data = localfs.read_file(local_path, cfg.local_roots())
    return await upload_bytes(graph, drive, folder, new_name or name, data, overwrite)


async def write_text_file(
    graph: Graph, drive: str | None, path: str, content: str, overwrite: bool = False
) -> dict:
    p = PurePosixPath(drives.clean_path(path))
    if p.suffix.lower() not in TEXT_EXTENSIONS:
        raise ValueError(f"write_text_file only writes text formats ({', '.join(sorted(TEXT_EXTENSIONS))}).")
    if not overwrite:
        base = await drives.drive_base(graph, drive)
        try:
            await graph.get(drives.item_ref(base, str(p)))
            raise ValueError(f"'{p}' already exists. Pass overwrite=true to replace it (OneDrive keeps the old version).")
        except GraphError as exc:
            if exc.status != 404:
                raise
    return await upload_bytes(graph, drive, str(p.parent), p.name, content.encode("utf-8"), overwrite=True)


async def create_folder(graph: Graph, drive: str | None, path: str) -> dict:
    base = await drives.drive_base(graph, drive)
    current = ""
    item: dict = {}
    for part in drives.clean_path(path).split("/"):
        parent = current
        current = f"{current}/{part}" if current else part
        try:
            item = await graph.get(drives.item_ref(base, current))
        except GraphError as exc:
            if exc.status != 404:
                raise
            item = await graph.post(
                drives.children_ref(base, parent),
                json={"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
            )
    return drives.describe(item)


async def move_item(
    graph: Graph, drive: str | None, path: str, destination_folder: str | None, new_name: str | None
) -> dict:
    if destination_folder is None and not new_name:
        raise ValueError("Give a destination_folder, a new_name, or both.")
    ref, item = await drives.get_item(graph, drive, path)
    body: dict = {}
    if new_name:
        body["name"] = new_name
    if destination_folder is not None:
        _, dest = await drives.get_item(graph, drive, destination_folder)
        if "folder" not in dest:
            raise ValueError(f"'{destination_folder}' is not a folder.")
        body["parentReference"] = {"id": dest["id"]}
    return drives.describe(await graph.patch(ref, json=body))


async def copy_item(
    graph: Graph, drive: str | None, path: str, destination_folder: str,
    destination_drive: str | None = None, new_name: str | None = None,
) -> dict:
    """Copy by download and re-upload so the copy exists when this returns."""
    ref, item = await drives.get_item(graph, drive, path)
    if "folder" in item:
        raise ValueError("Copying whole folders isn't supported; copy files one at a time.")
    data = await graph.get_bytes(f"{ref}/content")
    target_drive = destination_drive if destination_drive is not None else drive
    return await upload_bytes(graph, target_drive, destination_folder, new_name or item["name"], data)


MAX_DELETE = 25


async def items_for_deletion(graph: Graph, drive: str | None, paths: list[str]) -> list[tuple[str, dict]]:
    """Look up the files to delete. Empty folders are allowed; folders with anything in them are refused
    so one call can't remove a whole tree."""
    if not paths:
        raise ValueError("Give at least one file or empty folder path.")
    if len(paths) > MAX_DELETE:
        raise ValueError(f"Delete at most {MAX_DELETE} items at a time.")
    found = []
    for path in paths:
        ref, item = await drives.get_item(graph, drive, path)
        if "folder" in item and (item["folder"].get("childCount") or 0) != 0:
            raise ValueError(
                f"'{path}' is a folder with {item['folder'].get('childCount')} item(s) in it. Only files and empty "
                "folders can be deleted here; empty it first or delete it in OneDrive or SharePoint."
            )
        found.append((ref, item))
    return found


def describe_for_deletion(items: list[tuple[str, dict]]) -> str:
    n_folders = sum(1 for _, item in items if "folder" in item)
    n_files = len(items) - n_folders
    parts = []
    if n_files:
        parts.append(f"{n_files} file{'s' if n_files != 1 else ''}")
    if n_folders:
        parts.append(f"{n_folders} empty folder{'s' if n_folders != 1 else ''}")
    lines = [f"Delete {' and '.join(parts)}? They go to the recycle bin and can be restored."]
    for _, item in items:
        folder = (item.get("parentReference") or {}).get("path", "").split("root:", 1)[-1] or "/"
        detail = "empty folder" if "folder" in item else f"{item.get('size', 0):,} bytes"
        lines.append(f"• {item.get('name')}  (in {folder}, {detail})")
    return "\n".join(lines)


async def delete_items(graph: Graph, items: list[tuple[str, dict]]) -> list[dict]:
    results = []
    for ref, item in items:
        try:
            await graph.delete(ref)
            results.append({"name": item.get("name"), "deleted": True})
        except GraphError as exc:
            results.append({"name": item.get("name"), "deleted": False, "error": str(exc)})
    return results
