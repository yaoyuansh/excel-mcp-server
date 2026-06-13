"""
Excel MCP Server - Model Context Protocol server for Microsoft Excel.

Wraps the Excel Core engine (win32com COM API) as standard MCP tools,
allowing AI assistants to read, edit, and manage Excel workbooks through
natural language commands.

Transport: stdio (default), streamable HTTP
Requires: Windows + Microsoft Excel installed + running

Usage:
    # As MCP server (stdio transport)
    python -m excel_mcp.server

    # Or via entry point
    excel-mcp-server
"""

import json
import sys
import os
import asyncio
import logging
from typing import Any, Sequence

# MCP SDK imports
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    Tool,
    TextContent,
    EmbeddedResource,
)

# Configure logging to stderr (stdout is reserved for MCP protocol)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("excel-mcp")

# ─────────────────────────────────────────────────────────────
# Import Excel Core engine
# ─────────────────────────────────────────────────────────────
# Add parent directory to path for excel_core import
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from excel_core import (
        excel_read,
        excel_edit,
        excel_list_sheets,
        excel_manage_sheets,
        find_excel_workbook,
    )
    import pythoncom
    import win32com.client as win32

    EXCEL_CORE_AVAILABLE = True
    logger.info("Excel Core engine loaded successfully")
except ImportError as e:
    EXCEL_CORE_AVAILABLE = False
    logger.warning(f"Excel Core engine not available: {e}")
    logger.warning("Install dependencies: pip install pywin32 Pillow")


# ─────────────────────────────────────────────────────────────
# MCP Server Definition
# ─────────────────────────────────────────────────────────────
app = Server("excel-mcp-server")


def _make_result(data: Any) -> list[TextContent]:
    """Convert result dict to MCP TextContent."""
    text = json.dumps(data, ensure_ascii=False, default=str, indent=2)
    return [TextContent(type="text", text=text)]


def _make_error(message: str) -> list[TextContent]:
    """Create an error response."""
    return _make_result({"success": bool(0), "error": message})


# ─────────────────────────────────────────────────────────────
# Tool Definitions
# ─────────────────────────────────────────────────────────────

TOOLS = [
    Tool(
        name="list_workbooks",
        description=(
            "List all currently open Excel workbooks. "
            "Use this first to discover which files are available for editing."
        ),
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="list_sheets",
        description=(
            "List all worksheets in an open Excel workbook. "
            "Returns sheet names, visibility, and used range addresses. "
            "Use this to understand the workbook structure before reading data."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {
                    "type": "string",
                    "description": "Workbook name (e.g. 'data.xlsx'), full path, or partial match.",
                },
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="read_sheet",
        description=(
            "Read data from an Excel worksheet. Supports reading values, styles, "
            "merged cells, charts, tables, shapes, hyperlinks, data validations, "
            "conditional formats, filters, and print settings.\n\n"
            "IMPORTANT: Always call list_sheets first to get the used_range, "
            "then use 'range' parameter to read specific areas for large sheets.\n\n"
            "Parameters:\n"
            "- filename: Workbook name or path (required)\n"
            "- sheet: Worksheet name (optional, defaults to active sheet)\n"
            "- range: Cell range like 'A1:F20' (optional, defaults to UsedRange)\n"
            "- include_style: Include font/color/alignment/border info (default false)\n"
            "- max_rows: Max rows to return (default 500 without style, 200 with style)\n"
            "- max_cells: Max cells to return (default 8000 without style, 2000 with style)\n"
            "- head_rows/tail_rows: Return first N + last M rows for large sheets\n"
            "- include_charts: Return chart list (default false)\n"
            "- include_tables: Return ListObject tables (default false)\n"
            "- include_shapes: Return shapes/textboxes/images (default false)\n"
            "- include_hyperlinks: Return hyperlinks (default false)\n"
            "- include_data_validations: Return data validation rules (default false)\n"
            "- include_conditional_formats: Return conditional format rules (default false)\n"
            "- include_filters: Return auto filter state (default false)\n"
            "- include_print_settings: Return page/print setup (default false)"
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {"type": "string", "description": "Workbook name or path."},
                "sheet": {"type": "string", "description": "Worksheet name (optional)."},
                "range": {"type": "string", "description": "Cell range e.g. 'A1:F20' (optional)."},
                "include_style": {"type": "boolean", "description": "Include cell styles (default false)."},
                "max_rows": {"type": "integer", "description": "Max rows to return."},
                "max_cells": {"type": "integer", "description": "Max cells to return."},
                "head_rows": {"type": "integer", "description": "First N rows to return."},
                "tail_rows": {"type": "integer", "description": "Last M rows to return."},
                "include_charts": {"type": "boolean", "description": "Return chart list."},
                "include_tables": {"type": "boolean", "description": "Return ListObject tables."},
                "include_shapes": {"type": "boolean", "description": "Return shapes/textboxes."},
                "include_hyperlinks": {"type": "boolean", "description": "Return hyperlinks."},
                "include_data_validations": {"type": "boolean", "description": "Return data validations."},
                "include_conditional_formats": {"type": "boolean", "description": "Return conditional formats."},
                "include_filters": {"type": "boolean", "description": "Return auto filter state."},
                "include_print_settings": {"type": "boolean", "description": "Return print settings."},
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="edit_sheet",
        description=(
            "Edit an Excel workbook. This is the main editing tool with 90+ actions.\n\n"
            "Common actions (edit_action parameter):\n"
            "• VALUE & FORMULA: set_value, set_formula, convert_text_numbers\n"
            "• STYLE: set_style, number_format, apply_cell_style, copy_format\n"
            "• LAYOUT: merge, unmerge, insert_rows, delete_rows, insert_cols, delete_cols\n"
            "• DATA: sort, auto_filter, clear_filter, find_replace\n"
            "• CONDITIONAL FORMAT: conditional_format, conditional_format_v2, delete_conditional_format, list_conditional_formats\n"
            "• DATA VALIDATION: data_validation, data_validation_v2, circle_invalid_data\n"
            "• CHARTS: create_chart, list_charts, edit_chart, delete_chart, chart_style, export_chart\n"
            "• TABLES: list_tables, convert_to_table, table_style, table_total_row, table_resize, convert_to_range\n"
            "• DATA OPS: remove_duplicates, text_to_columns, fill_series, auto_fill, subtotal, group, ungroup\n"
            "• SHAPES: list_shapes, insert_shape, insert_textbox, delete_shape, edit_shape, replace_image, move_shape\n"
            "• PAGE/PRINT: page_setup, print_area, print_titles, page_break, set_header, set_footer, print_preview\n"
            "• SHEET: freeze_panes, unfreeze_panes, hide_rows, show_rows, hide_cols, show_cols\n"
            "• APPEARANCE: set_tab_color, set_zoom, set_gridlines_visible, set_headings_visible, move_sheet, split\n"
            "• PROTECTION: protect, unprotect\n"
            "• WORKBOOK: set_workbook_property, trace_precedents, trace_dependents, clear_arrows\n"
            "• OTHER: add_comment, delete_comment, add_hyperlink, clear_contents, clear_all, copy_paste, insert_image, auto_fit, named_range, pivot_table, calculate\n\n"
            "Each action has its own parameters. See docs/TOOLS.md for full details."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {"type": "string", "description": "Workbook name or path."},
                "sheet": {"type": "string", "description": "Worksheet name (optional)."},
                "edit_action": {
                    "type": "string",
                    "description": "Action to perform (e.g. 'set_value', 'set_formula', 'set_style', 'sort', 'create_chart', etc.).",
                },
                "operations": {
                    "type": "array",
                    "items": {"type": "object"},
                    "description": "Array of operation objects. Each action has its own operation format.",
                },
                "cell": {"type": "string", "description": "Cell reference (e.g. 'A1')."},
                "value": {"description": "Value to write."},
                "formula": {"type": "string", "description": "Excel formula (e.g. '=SUM(A1:A10)')."},
                "range": {"type": "string", "description": "Cell range (e.g. 'A1:F20')."},
                "style": {"type": "object", "description": "Style object for set_style."},
                "row": {"type": "integer", "description": "Row number."},
                "col": {"type": "string", "description": "Column letter."},
                "count": {"type": "integer", "description": "Count for insert/delete operations."},
                "sort_column": {"type": "integer", "description": "Column number to sort by (1-based)."},
                "order": {"type": "string", "description": "Sort order: 'asc' or 'desc'."},
                "has_header": {"type": "boolean", "description": "Whether data has headers."},
                "field": {"type": "integer", "description": "Filter field number (1-based)."},
                "criteria": {"description": "Filter criteria (string, number, or array)."},
                "cf_type": {"type": "string", "description": "Conditional format type."},
                "operator": {"type": "integer", "description": "Comparison operator (1-8)."},
                "formula1": {"type": "string", "description": "First formula/value for conditional format."},
                "formula2": {"type": "string", "description": "Second formula/value (for between)."},
                "format": {"type": "object", "description": "Format object for conditional format."},
                "chart_type": {"type": "string", "description": "Chart type (column/bar/line/pie/scatter/area/doughnut/radar)."},
                "title": {"type": "string", "description": "Chart or element title."},
                "position": {"type": "string", "description": "Position for chart/element."},
                "width": {"type": "number", "description": "Width in pixels."},
                "height": {"type": "number", "description": "Height in pixels."},
                "path": {"type": "string", "description": "File path for images."},
                "text": {"type": "string", "description": "Text content for comments, hyperlinks, etc."},
                "url": {"type": "string", "description": "URL for hyperlinks."},
                "name": {"type": "string", "description": "Name for sheets, shapes, tables, etc."},
                "nr_action": {"type": "string", "description": "Named range action: create/delete/list."},
                "format_string": {"type": "string", "description": "Number format string."},
                "password": {"type": "string", "description": "Protection password."},
                "rows": {"type": "string", "description": "Row range e.g. '5:10'."},
                "cols": {"type": "string", "description": "Column range e.g. 'C:E'."},
                "cell_style": {"type": "string", "description": "Built-in cell style name."},
                "color": {"type": "string", "description": "Color value."},
                "color1": {"type": "string", "description": "Gradient start color."},
                "color2": {"type": "string", "description": "Gradient end color."},
                "direction": {"type": "string", "description": "Gradient direction."},
                "pattern": {"type": "string", "description": "Fill pattern."},
                "tint": {"type": "number", "description": "Tint/shade value (-1.0 to 1.0)."},
                "row_height": {"type": "number", "description": "Row height."},
                "column_width": {"type": "number", "description": "Column width."},
                "source": {"type": "string", "description": "Source range for copy_paste."},
                "destination": {"type": "string", "description": "Destination for copy_paste."},
                "fit_type": {"type": "string", "description": "Auto fit type: columns/rows/both."},
                "table_name": {"type": "string", "description": "ListObject table name."},
                "table_index": {"type": "integer", "description": "Table index (1-based)."},
                "chart_index": {"type": "integer", "description": "Chart index (1-based)."},
                "chart_name": {"type": "string", "description": "Chart shape name."},
                "chart_title": {"type": "string", "description": "Chart title for lookup."},
                "series_format": {"type": "array", "items": {"type": "object"}, "description": "Chart series formatting."},
                "trendlines": {"type": "array", "items": {"type": "object"}, "description": "Trendline config."},
                "x_axis": {"type": "object", "description": "X-axis config."},
                "y_axis": {"type": "object", "description": "Y-axis config."},
                "gridlines": {"type": "object", "description": "Gridline config."},
                "show_data_labels": {"type": "boolean", "description": "Show data labels on chart."},
                "show_legend": {"type": "boolean", "description": "Show chart legend."},
                "sort_keys": {"type": "array", "items": {"type": "object"}, "description": "Multi-column sort keys."},
                "validation_type": {"type": "string", "description": "Data validation type."},
                "items": {"type": "array", "items": {"type": "string"}, "description": "List items for data validation."},
                "alert_style": {"type": "string", "description": "Validation alert style."},
                "input_title": {"type": "string", "description": "Validation input title."},
                "input_message": {"type": "string", "description": "Validation input message."},
                "error_title": {"type": "string", "description": "Validation error title."},
                "error_message": {"type": "string", "description": "Validation error message."},
                "shape_type": {"type": "string", "description": "Shape type (rectangle/oval/triangle etc.)."},
                "find": {"type": "string", "description": "Find text for find_replace."},
                "replace": {"type": "string", "description": "Replace text for find_replace."},
                "setup": {"type": "object", "description": "Page setup config."},
                "orientation": {"type": "string", "description": "Page orientation: portrait/landscape."},
                "paper_size": {"type": "string", "description": "Paper size: a4/a3/letter/legal."},
                "zoom": {"type": "integer", "description": "Zoom percentage."},
                "visible": {"type": "boolean", "description": "Visibility flag."},
                "operation": {"type": "string", "description": "Sub-operation type."},
            },
            "required": ["filename", "edit_action"],
        },
    ),
    Tool(
        name="manage_sheets",
        description=(
            "Manage worksheets in an Excel workbook.\n\n"
            "Actions (sheet_action parameter):\n"
            "- add: Create new worksheet (optionally with 'name')\n"
            "- delete: Delete worksheet by 'name'\n"
            "- rename: Rename worksheet ('old_name' → 'new_name')\n"
            "- copy: Copy worksheet ('name', optionally 'after' another sheet)\n"
            "- hide: Hide worksheet by 'name'\n"
            "- show: Show hidden worksheet by 'name'\n"
            "- activate: Switch to worksheet by 'name'"
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {"type": "string", "description": "Workbook name or path."},
                "sheet_action": {
                    "type": "string",
                    "description": "Action: add, delete, rename, copy, hide, show, activate.",
                    "enum": ["add", "delete", "rename", "copy", "hide", "show", "activate"],
                },
                "name": {"type": "string", "description": "Worksheet name."},
                "old_name": {"type": "string", "description": "Current worksheet name (for rename)."},
                "new_name": {"type": "string", "description": "New worksheet name (for rename)."},
                "after": {"type": "string", "description": "Copy after this sheet name."},
            },
            "required": ["filename", "sheet_action"],
        },
    ),
]


# ─────────────────────────────────────────────────────────────
# MCP Handlers
# ─────────────────────────────────────────────────────────────

@app.list_tools()
async def list_tools() -> list[Tool]:
    """Return available tools."""
    return TOOLS


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    """Handle tool calls."""
    if not EXCEL_CORE_AVAILABLE:
        return _make_error(
            "Excel Core engine not available. "
            "Ensure pywin32 and Pillow are installed: pip install pywin32 Pillow"
        )

    try:
        # Initialize COM for this call
        pythoncom.CoInitialize()

        # Dispatch to the appropriate handler
        handler = TOOL_HANDLERS.get(name)
        if not handler:
            return _make_error(f"Unknown tool: {name}")

        result = handler(arguments)
        return _make_result(result)

    except Exception as e:
        logger.exception(f"Error in tool '{name}'")
        return _make_error(f"{type(e).__name__}: {str(e)}")

    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────
# Tool Handlers
# ─────────────────────────────────────────────────────────────

def handle_list_workbooks(params: dict) -> dict:
    """List all open Excel workbooks."""
    try:
        excel = win32.GetActiveObject("Excel.Application")
        workbooks = []
        for wb in excel.Workbooks:
            workbooks.append({
                "name": wb.Name,
                "full_path": wb.FullName,
                "sheet_count": wb.Sheets.Count,
            })
        return {"success": bool(1), "workbooks": workbooks, "count": len(workbooks)}
    except Exception as e:
        return {"success": bool(0), "error": f"Cannot connect to Excel: {e}. Is Excel running?"}


def handle_list_sheets(params: dict) -> dict:
    """List worksheets in a workbook."""
    return excel_list_sheets(params)


def handle_read_sheet(params: dict) -> dict:
    """Read data from a worksheet."""
    return excel_read(params)


def handle_edit_sheet(params: dict) -> dict:
    """Edit a worksheet."""
    return excel_edit(params)


def handle_manage_sheets(params: dict) -> dict:
    """Manage worksheets."""
    return excel_manage_sheets(params)


TOOL_HANDLERS = {
    "list_workbooks": handle_list_workbooks,
    "list_sheets": handle_list_sheets,
    "read_sheet": handle_read_sheet,
    "edit_sheet": handle_edit_sheet,
    "manage_sheets": handle_manage_sheets,
}


# ─────────────────────────────────────────────────────────────
# Server Entry Point
# ─────────────────────────────────────────────────────────────

async def run_stdio():
    """Run MCP server with stdio transport."""
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


def main():
    """Entry point for the MCP server."""
    logger.info("Starting Excel MCP Server (stdio transport)...")
    asyncio.run(run_stdio())


if __name__ == "__main__":
    main()
