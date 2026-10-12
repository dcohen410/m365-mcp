import os
import tempfile

# Keep tests away from any real settings or saved sign-in.
os.environ["M365_MCP_CONFIG_DIR"] = tempfile.mkdtemp(prefix="m365mcp-test-")

import httpx
import pytest
import respx

from m365_mcp.config import Config
from m365_mcp.graph import GRAPH, Graph


@pytest.fixture
def mock():
    with respx.mock(base_url=GRAPH, assert_all_called=False) as router:
        yield router


@pytest.fixture
async def graph():
    g = Graph(lambda: "test-token", httpx.AsyncClient(follow_redirects=True))
    yield g
    await g.aclose()


@pytest.fixture
def cfg(tmp_path):
    return Config(client_id="cid", tenant_id="tid", local_dirs=[str(tmp_path)])


def item(name, item_id="ITEM", drive_id="DRIVE", parent_id="PARENT", folder=False, path="/drive/root:/Docs"):
    out = {
        "id": item_id, "name": name, "size": 10, "webUrl": f"https://example/{name}",
        "parentReference": {"driveId": drive_id, "id": parent_id, "path": path},
    }
    if folder:
        out["folder"] = {"childCount": 0}
    else:
        out["file"] = {}
    return out
