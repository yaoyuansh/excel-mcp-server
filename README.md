# Excel MCP Server

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.10+-green.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)]()

> 🇨🇳 [中文文档](README_CN.md) | 🇺🇸 English

A **Model Context Protocol (MCP) server** for Microsoft Excel automation. Control Excel directly through AI assistants using natural language — read cells, write formulas, apply formatting, create charts, manage sheets, and **90+ more operations**.

## Features

### Read & Write
- Read cell values, formulas, styles, merged cells
- Set values and formulas with auto type detection
- Copy, paste, find & replace

### Formatting & Styles
- Font, color, alignment, borders, number formats
- Conditional formatting (cell value, formula, databar, color scale, icon set, top/bottom, duplicates, etc.)
- Cell styles (gradient fill, pattern fill, tint/shade, diagonal borders, built-in styles)
- Row height, column width, format copying

### Data Operations
- Sort (single & multi-column), auto filter
- Remove duplicates, text-to-columns, fill series
- Subtotals, grouping/ungrouping, outline levels
- Data validation (list, number, date, time, text length, custom formula)

### Charts
- Create, edit, delete charts (column, bar, line, pie, scatter, area, doughnut, radar)
- Deep editing: trendlines, single-point coloring, data labels, axis config, gridlines
- Export chart as image

### Tables (ListObject)
- Create, style, resize Excel tables
- Total row with custom aggregation functions
- Convert tables to/from ranges

### Shapes & Textboxes
- Insert, edit, move, delete shapes (rectangle, oval, triangle, star, arrow, etc.)
- Insert and edit textboxes with full formatting

### Page Layout & Print
- Page setup (orientation, paper size, margins, zoom, fit-to-pages)
- Print area, print titles, page breaks
- Headers & footers with dynamic fields (date, page number, sheet name)

### Sheet Management
- Add, delete, rename, copy, hide, show, activate worksheets
- Tab color, zoom, gridlines, headings visibility
- Split windows, freeze panes, calculation mode

### Workbook Properties
- Set document properties (title, author, keywords, etc.)
- Formula auditing (trace precedents, trace dependents)
- Named ranges, comments, hyperlinks, protection

## Quick Start

### Prerequisites

- **Windows** with **Microsoft Excel** installed and running
- **Python 3.10+**
- An MCP-compatible AI client (Claude Desktop, Cursor, etc.)

### Install

```bash
git clone https://github.com/yaoyuansh/excel-mcp-server.git
cd excel-mcp-server
pip install -e .
```

### Configure MCP Client

**Claude Code**:

```bash
claude mcp add excel-mcp-server -s user -- uv run --directory C:/path/to/excel-mcp-server excel-mcp-server
```

Or use the equivalent JSON configuration:

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "type": "stdio",
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "C:/path/to/excel-mcp-server",
        "excel-mcp-server"
      ]
    }
  }
}
```

**Claude Desktop** (`claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "python",
      "args": ["-m", "excel_mcp.server"],
      "cwd": "C:/path/to/excel-mcp-server/src"
    }
  }
}
```

**Cursor** (`.cursor/mcp.json`):
```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "python",
      "args": ["-m", "excel_mcp.server"],
      "cwd": "C:/path/to/excel-mcp-server/src"
    }
  }
}
```

## Tools

| Tool | Description |
|------|-------------|
| `list_workbooks` | List all open Excel workbooks |
| `list_sheets` | List worksheets with used ranges |
| `read_sheet` | Read data, styles, charts, tables, shapes, validations, filters, print settings |
| `edit_sheet` | Edit with 90+ actions (values, formulas, formatting, charts, tables, shapes, page layout, etc.) |
| `manage_sheets` | Add, delete, rename, copy, hide, show, activate worksheets |

See [docs/TOOLS.md](docs/TOOLS.md) for full API reference.

## Architecture

```
┌─────────────────┐     stdio/HTTP      ┌──────────────────┐
│  AI Assistant   │ ◄──────────────────► │  MCP Server      │
│  (Claude/Cursor)│                      │  (server.py)     │
└─────────────────┘                      └────────┬─────────┘
                                                  │
                                          ┌───────▼─────────┐
                                          │  Excel Core      │
                                          │  (excel_core.py) │
                                          │  win32com COM    │
                                          └───────┬─────────┘
                                                  │ COM API
                                          ┌───────▼─────────┐
                                          │  Microsoft Excel │
                                          │  (running app)   │
                                          └─────────────────┘
```

## Project Structure

```
excel-mcp-server/
├── .gitignore              # Git ignore rules
├── README.md              # English documentation
├── README_CN.md           # 中文文档
├── LICENSE                # Apache 2.0
├── pyproject.toml         # Python packaging config
├── docs/
│   └── TOOLS.md           # Full API reference (EN + CN)
├── tests/                 # Regression and metadata tests
└── src/
    └── excel_mcp/
        ├── __init__.py    # Package init
        ├── server.py      # MCP server (5 tools)
        └── excel_core.py  # Excel operations engine (90+ actions)
```

## Development

Run the regression test suite:

```bash
pip install -e ".[dev]"
pytest
```

The tests cover package import, MCP tool schemas, license metadata consistency, and protection against accidental lowercase JSON literals (`true` / `false` / `null`) in Python code.

## License

Apache 2.0 — see [LICENSE](LICENSE)

Copyright 2026 yaoyuansh
