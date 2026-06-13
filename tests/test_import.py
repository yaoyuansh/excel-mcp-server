"""Import smoke tests for excel-mcp-server."""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


def test_server_imports_and_exposes_tools():
    import excel_mcp.server as server

    assert server.EXCEL_CORE_AVAILABLE is bool(1)
    assert len(server.TOOLS) == 5
    assert [tool.name for tool in server.TOOLS] == [
        "list_workbooks",
        "list_sheets",
        "read_sheet",
        "edit_sheet",
        "manage_sheets",
    ]


def test_make_error_returns_json_text_content():
    import json
    import excel_mcp.server as server

    payload = json.loads(server._make_error("boom")[0].text)
    assert payload == {"success": bool(0), "error": "boom"}
