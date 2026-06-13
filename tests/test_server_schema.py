"""MCP tool schema tests."""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import excel_mcp.server as server


def _tool_map():
    return {tool.name: tool for tool in server.TOOLS}


def test_expected_tool_names_are_registered():
    assert set(_tool_map()) == {
        "list_workbooks",
        "list_sheets",
        "read_sheet",
        "edit_sheet",
        "manage_sheets",
    }


def test_tool_schemas_are_json_schema_objects():
    for tool in server.TOOLS:
        schema = tool.inputSchema
        assert schema["type"] == "object"
        assert "properties" in schema
        assert "required" in schema


def test_required_arguments_for_stateful_tools():
    tools = _tool_map()
    assert tools["list_workbooks"].inputSchema["required"] == []
    assert tools["list_sheets"].inputSchema["required"] == ["filename"]
    assert tools["read_sheet"].inputSchema["required"] == ["filename"]
    assert tools["edit_sheet"].inputSchema["required"] == ["filename", "edit_action"]
    assert tools["manage_sheets"].inputSchema["required"] == ["filename", "sheet_action"]


def test_edit_sheet_has_core_parameters():
    props = _tool_map()["edit_sheet"].inputSchema["properties"]
    for name in [
        "filename",
        "sheet",
        "edit_action",
        "operations",
        "range",
        "cell",
        "chart_type",
        "table_name",
        "cf_type",
        "validation_type",
        "shape_type",
        "sort_keys",
    ]:
        assert name in props
