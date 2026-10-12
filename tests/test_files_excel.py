import json

import pytest

from m365_mcp import drives, excel, files, localfs
from m365_mcp.graph import SIMPLE_UPLOAD_LIMIT
from tests.conftest import item


async def test_drive_from_sharepoint_url(mock, graph):
    route = mock.get("/sites/contoso.sharepoint.com:/sites/Finance").respond(json={"id": "SITE1"})
    assert await drives.drive_base(graph, "https://contoso.sharepoint.com/sites/Finance/Shared Documents") == "/sites/SITE1/drive"
    assert await drives.drive_base(graph, None) == "/me/drive"
    assert await drives.drive_base(graph, "b!abc") == "/drives/b!abc"
    assert route.call_count == 1


def test_item_paths_are_quoted():
    assert drives.item_ref("/me/drive", "/Staff Records/Jane/") == "/me/drive/root:/Staff%20Records/Jane:"
    assert drives.item_ref("/me/drive", "") == "/me/drive/root"
    assert drives.children_ref("/me/drive", "A\\B") == "/me/drive/root:/A/B:/children"


def test_local_paths_are_limited(tmp_path):
    roots = [tmp_path.resolve()]
    with pytest.raises(localfs.LocalPathError):
        localfs.write_file(str(tmp_path / ".." / "escape"), "x.txt", b"x", roots)
    saved = localfs.write_file(str(tmp_path / "sub"), "a/b:c.txt", b"x", roots)
    assert saved.name == "a_b_c.txt"


async def test_large_upload_uses_session(mock, graph):
    mock.post("/me/drive/root:/Big/file.bin:/createUploadSession").respond(json={"uploadUrl": "https://up.example/s"})
    chunks = mock.put("https://up.example/s").respond(json=item("file.bin"))
    data = b"x" * (SIMPLE_UPLOAD_LIMIT + 1)
    await files.upload_bytes(graph, None, "Big", "file.bin", data)
    ranges = [c.request.headers["Content-Range"] for c in chunks.calls]
    assert ranges[0].startswith("bytes 0-") and ranges[-1].endswith(f"/{len(data)}")


async def test_write_text_file_refuses_to_overwrite(mock, graph):
    mock.get("/me/drive/root:/notes.md:").respond(json=item("notes.md"))
    with pytest.raises(ValueError, match="already exists"):
        await files.write_text_file(graph, None, "notes.md", "hi")


async def test_write_text_file_creates_new(mock, graph):
    mock.get("/me/drive/root:/notes.md:").respond(status_code=404, json={"error": {"code": "itemNotFound", "message": "x"}})
    put = mock.put("/me/drive/root:/notes.md:/content").respond(json=item("notes.md"))
    await files.write_text_file(graph, None, "notes.md", "hi")
    assert put.calls[0].request.content == b"hi"


async def test_read_word_doc_via_pdf_conversion(mock, graph, monkeypatch):
    mock.get("/me/drive/root:/a.docx:").respond(json=item("a.docx"))
    conv = mock.get("/drives/DRIVE/items/ITEM/content").respond(content=b"%PDF")
    monkeypatch.setattr(files, "pdf_text", lambda data, n: "text from pdf")
    assert await files.read_file(graph, None, "a.docx") == "text from pdf"
    assert conv.calls[0].request.url.params["format"] == "pdf"


async def test_retry_after_429(mock, graph, monkeypatch):
    async def no_sleep(_):
        return None
    monkeypatch.setattr("asyncio.sleep", no_sleep)
    route = mock.get("/me/drive").mock(side_effect=[
        __import__("httpx").Response(429, headers={"Retry-After": "1"}),
        __import__("httpx").Response(200, json={"id": "D"}),
    ])
    assert (await graph.get("/me/drive"))["id"] == "D"
    assert route.call_count == 2


def test_column_maths():
    assert excel.col_to_num("A") == 1 and excel.col_to_num("AA") == 27
    assert excel.num_to_col(28) == "AB"
    assert excel._range_for("B", 5, [[1, 2, 3], [4, 5, 6]]) == "B5:D6"
    with pytest.raises(ValueError):
        excel._range_for("A", 1, [[1], [1, 2]])


async def test_excel_append_backs_up_then_writes_below_used_range(mock, graph):
    mock.get("/me/drive/root:/Register.xlsx:").respond(json=item("Register.xlsx"))
    mock.get("/drives/DRIVE/items/ITEM/content").respond(content=b"xlsx-bytes")
    backup = mock.put(url__regex=r".*/drives/DRIVE/items/PARENT:/Register%20%28backup%20.*%29\.xlsx:/content").respond(
        json={"name": "Register (backup).xlsx"})
    mock.get(url__regex=r".*/workbook/worksheets\('Team%20Register'\)/usedRange.*").respond(
        json={"address": "'Team Register'!A1:F51"})
    write = mock.patch(url__regex=r".*/workbook/worksheets\('Team%20Register'\)/range\(address='A52:B52'\)").respond(json={})
    out = await excel.append_rows(graph, None, "Register.xlsx", [["Jane", "Active"]], sheet="Team Register")
    assert backup.called and write.called
    assert json.loads(write.calls[0].request.content) == {"values": [["Jane", "Active"]]}
    assert out["range"] == "Team Register!A52:B52"


async def test_excel_rejects_non_workbooks(mock, graph):
    mock.get("/me/drive/root:/a.csv:").respond(json=item("a.csv"))
    with pytest.raises(ValueError):
        await excel.list_sheets(graph, None, "a.csv")
