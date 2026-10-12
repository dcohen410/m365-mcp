"""Reading and editing Excel workbooks stored in OneDrive or SharePoint."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import PurePosixPath
from urllib.parse import quote

from . import drives
from .graph import Graph

_ADDRESS = re.compile(r"^(?:.*!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$")


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _sheet(name: str) -> str:
    return f"worksheets('{quote(name.replace(chr(39), chr(39) * 2))}')"


async def _workbook(graph: Graph, drive: str | None, path: str) -> tuple[str, dict]:
    ref, item = await drives.get_item(graph, drive, path)
    if PurePosixPath(item["name"]).suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError("Only .xlsx and .xlsm workbooks can be read or edited.")
    return ref, item


async def backup(graph: Graph, ref: str, item: dict) -> str:
    """Save a dated copy next to the workbook before changing it."""
    p = PurePosixPath(item["name"])
    name = f"{p.stem} (backup {datetime.now():%Y-%m-%d %H%M%S}){p.suffix}"
    data = await graph.get_bytes(f"{ref}/content")
    drive_id = item["parentReference"]["driveId"]
    parent_id = item["parentReference"]["id"]
    resp = await graph.request(
        "PUT",
        f"/drives/{drive_id}/items/{parent_id}:/{quote(name)}:/content",
        params={"@microsoft.graph.conflictBehavior": "rename"},
        content=data,
        headers={"Content-Type": "application/octet-stream"},
    )
    return resp.json().get("name", name)


async def list_sheets(graph: Graph, drive: str | None, path: str) -> dict:
    ref, _ = await _workbook(graph, drive, path)
    sheets = await graph.get_all(f"{ref}/workbook/worksheets")
    tables = await graph.get_all(f"{ref}/workbook/tables")
    return {
        "sheets": [s["name"] for s in sheets],
        "tables": [{"name": t["name"]} for t in tables],
    }


async def read_range(graph: Graph, drive: str | None, path: str, sheet: str, address: str | None) -> dict:
    ref, _ = await _workbook(graph, drive, path)
    if address:
        url = f"{ref}/workbook/{_sheet(sheet)}/range(address='{quote(address)}')"
    else:
        url = f"{ref}/workbook/{_sheet(sheet)}/usedRange(valuesOnly=true)"
    rng = await graph.get(url, params={"$select": "address,values,rowCount,columnCount"})
    return {"address": rng.get("address"), "values": rng.get("values", [])}


def _range_for(start_col: str, start_row: int, values: list[list]) -> str:
    if not values or not values[0]:
        raise ValueError("values must be a non-empty list of rows.")
    width = len(values[0])
    if any(len(r) != width for r in values):
        raise ValueError("Every row in values must have the same number of cells.")
    end_col = num_to_col(col_to_num(start_col) + width - 1)
    return f"{start_col}{start_row}:{end_col}{start_row + len(values) - 1}"


async def write_range(
    graph: Graph, drive: str | None, path: str, sheet: str, start_cell: str, values: list[list],
    make_backup: bool = True,
) -> dict:
    ref, item = await _workbook(graph, drive, path)
    m = _ADDRESS.match(start_cell.upper().strip())
    if not m:
        raise ValueError(f"'{start_cell}' isn't a cell address like B2.")
    address = _range_for(m.group(1), int(m.group(2)), values)
    backup_name = await backup(graph, ref, item) if make_backup else None
    rng = await graph.patch(f"{ref}/workbook/{_sheet(sheet)}/range(address='{address}')", json={"values": values})
    return {"updated": rng.get("address", address), "backup": backup_name}


async def append_rows(
    graph: Graph, drive: str | None, path: str, values: list[list], sheet: str | None = None,
    table: str | None = None, make_backup: bool = True,
) -> dict:
    if not (sheet or table):
        raise ValueError("Give either a sheet or a table to append to.")
    ref, item = await _workbook(graph, drive, path)
    backup_name = await backup(graph, ref, item) if make_backup else None
    if table:
        await graph.post(f"{ref}/workbook/tables('{quote(table)}')/rows", json={"values": values})
        return {"appended_rows": len(values), "table": table, "backup": backup_name}
    used = await graph.get(f"{ref}/workbook/{_sheet(sheet)}/usedRange(valuesOnly=true)", params={"$select": "address"})
    m = _ADDRESS.match(used["address"])
    if not m:
        raise ValueError(f"Couldn't understand the sheet's used range {used['address']}.")
    start_col, last_row = m.group(1), int(m.group(4) or m.group(2))
    address = _range_for(start_col, last_row + 1, values)
    await graph.patch(f"{ref}/workbook/{_sheet(sheet)}/range(address='{address}')", json={"values": values})
    return {"appended_rows": len(values), "range": f"{sheet}!{address}", "backup": backup_name}
