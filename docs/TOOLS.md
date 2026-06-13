# Excel MCP Server - API Reference

> 🇨🇳 [中文](#中文参考) | 🇺🇸 [English](#english-reference)

---

## English Reference

### Overview

The Excel MCP Server exposes 5 tools via the Model Context Protocol. All operations work on **currently open** Excel workbooks (via COM API).

---

### Tool 1: `list_workbooks`

List all currently open Excel workbooks.

**Parameters:** null

**Response:**
```json
{
  "success": true,
  "workbooks": [
    { "name": "data.xlsx", "sheets": 3, "path": "C:\\Users\\..." }
  ]
}
```

---

### Tool 2: `list_sheets`

List all worksheets in a workbook with their used range addresses.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `filename` | string | ✅ | Workbook name, full path, or partial match |

**Response:**
```json
{
  "success": true,
  "workbook": "data.xlsx",
  "sheets": [
    { "index": 1, "name": "Sheet1", "visible": true, "used_range": "$A$1:$F$100" }
  ]
}
```

---

### Tool 3: `read_sheet`

Read data from an Excel worksheet. Supports reading values, styles, charts, tables, shapes, hyperlinks, data validations, conditional formats, filters, and print settings.

**⚠️ Always call `list_sheets` first to get the `used_range`, then use `range` to read specific areas for large sheets.**

**Parameters:**

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `filename` | string | ✅ | - | Workbook name or path |
| `sheet` | string | ❌ | Active sheet | Worksheet name |
| `range` | string | ❌ | UsedRange | Cell range like `A1:F20` |
| `include_style` | boolean | ❌ | false | Include font/color/alignment/border info |
| `max_rows` | integer | ❌ | 500/200 | Max rows (200 with style, 500 without) |
| `max_cells` | integer | ❌ | 8000/2000 | Max cells |
| `head_rows` | integer | ❌ | - | First N rows (for large sheets) |
| `tail_rows` | integer | ❌ | - | Last M rows |
| `include_charts` | boolean | ❌ | false | Return chart list |
| `include_tables` | boolean | ❌ | false | Return ListObject tables |
| `include_shapes` | boolean | ❌ | false | Return shapes/textboxes |
| `include_hyperlinks` | boolean | ❌ | false | Return hyperlinks |
| `include_data_validations` | boolean | ❌ | false | Return data validations |
| `include_conditional_formats` | boolean | ❌ | false | Return conditional formats |
| `include_filters` | boolean | ❌ | false | Return auto filter state |
| `include_print_settings` | boolean | ❌ | false | Return print settings |

**Example:**
```json
{
  "filename": "report.xlsx",
  "sheet": "Sales",
  "range": "A1:F50",
  "include_style": true,
  "include_conditional_formats": true
}
```

---

### Tool 4: `edit_sheet`

The main editing tool with **90+ actions**. All actions operate on the currently open workbook.

**Common Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `filename` | string | **Required.** Workbook name or path |
| `edit_action` | string | **Required.** Action type (see below) |
| `sheet` | string | Worksheet name (defaults to active sheet) |
| `operations` | array | Array of operation objects (for batch operations) |
| `range` | string | Cell range like `A1:F20` |
| `cell` | string | Single cell like `G3` |
| `row` | integer | Row number (1-based) |
| `col` | string | Column letter like `C` |
| `count` | integer | Count for insert/delete |
| `style` | object | Style object (see Style section) |
| `text` | string | Text content |
| `path` | string | File path for images |
| `name` | string | Name for sheets, shapes, tables |
| `width` | number | Width in pixels |
| `height` | number | Height in pixels |

---

#### 4.1 Value & Formula

**`set_value`** — Set cell value
```json
{
  "filename": "data.xlsx",
  "edit_action": "set_value",
  "operations": [
    {"cell": "A1", "value": "Hello"},
    {"cell": "B1", "value": 42}
  ]
}
```

**`set_formula`** — Set cell formula
```json
{
  "filename": "data.xlsx",
  "edit_action": "set_formula",
  "operations": [
    {"cell": "C1", "formula": "=SUM(A1:B1)"},
    {"cell": "D1", "formula": "=AVERAGE(A1:B1)"}
  ]
}
```

**`convert_text_numbers`** — Convert text-formatted numbers to real numbers
```json
{
  "filename": "data.xlsx",
  "edit_action": "convert_text_numbers",
  "range": "C4:C100"
}
```

---

#### 4.2 Style & Format

**`set_style`** — Apply styles to cells
```json
{
  "filename": "data.xlsx",
  "edit_action": "set_style",
  "operations": [
    {
      "range": "A1:F1",
      "style": {
        "font": {"name": "微软雅黑", "size": 12, "bold": true, "color": "红"},
        "h_align": "center",
        "v_align": "center",
        "fill_color": "黄",
        "borders": {"all": {"style": "continuous", "weight": "thin", "color": "黑"}}
      }
    }
  ]
}
```

**Style Object Properties:**
- `font`: `{name, size, bold, italic, color, underline, strikethrough}`
- `h_align`: `left`, `center`, `right`, `justify` (or Chinese: 左/居中/右)
- `v_align`: `top`, `center`, `bottom`
- `wrap_text`: boolean
- `fill_color`: color name or hex
- `number_format`: format string (e.g., `"¥#,##0.00"`, `"0.00%"`)
- `borders`: `{all: {style, weight, color}}` or `{top: {...}, bottom: {...}, left: {...}, right: {...}}`
- `orientation`: degrees (0-360)
- `indent`: integer
- `column_width`: number
- `row_height`: number
- `auto_fit_column`: boolean

**`number_format`** — Set number format
```json
{"filename": "data.xlsx", "edit_action": "number_format", "range": "C1:C20", "format_string": "¥#,##0.00"}
```
Format shortcuts: `currency` (¥#,##0.00), `percent` (0.00%), `date` (yyyy-mm-dd)

**`apply_cell_style`** — Apply built-in Excel style
```json
{"filename": "data.xlsx", "edit_action": "apply_cell_style", "range": "A1:F1", "cell_style": "heading_1"}
```
Styles: `good`, `bad`, `neutral`, `heading_1`, `title`, `accent_1`, `currency`, `percent`, etc.

**`copy_format`** — Copy formatting (format painter)
```json
{"filename": "data.xlsx", "edit_action": "copy_format", "source": "A1:F1", "destination": "A2:F100"}
```

**`gradient_fill`** — Gradient fill
```json
{"filename": "data.xlsx", "edit_action": "gradient_fill", "range": "A1:F1", "color1": "红", "color2": "白", "direction": "horizontal"}
```

**`pattern_fill`** — Pattern fill
```json
{"filename": "data.xlsx", "edit_action": "pattern_fill", "range": "A1:F1", "pattern": "checker", "fill_color": "黄", "pattern_color": "红"}
```

**`tint_and_shade`** — Adjust color brightness
```json
{"filename": "data.xlsx", "edit_action": "tint_and_shade", "range": "A1:F1", "tint": 0.4, "apply_to": "fill"}
```

**`diagonal_borders`** — Diagonal borders (for header cells)
```json
{"filename": "data.xlsx", "edit_action": "diagonal_borders", "range": "A1", "direction": "down", "weight": "thin", "color": "黑"}
```

**`set_row_height`** / **`set_column_width`**
```json
{"filename": "data.xlsx", "edit_action": "set_row_height", "height": 25, "rows": "5:10"}
{"filename": "data.xlsx", "edit_action": "set_column_width", "width": 15, "cols": "C:E"}
```

---

#### 4.3 Layout

**`merge`** / **`unmerge`**
```json
{"filename": "data.xlsx", "edit_action": "merge", "range": "A1:F1"}
{"filename": "data.xlsx", "edit_action": "unmerge", "range": "A1:F1"}
```

**`insert_rows`** / **`delete_rows`** / **`insert_cols`** / **`delete_cols`**
```json
{"filename": "data.xlsx", "edit_action": "insert_rows", "row": 5, "count": 3}
{"filename": "data.xlsx", "edit_action": "delete_cols", "col": "D", "count": 2}
```

---

#### 4.4 Data Operations

**`sort`** — Sort data (single or multi-column)
```json
{"filename": "data.xlsx", "edit_action": "sort", "range": "A1:J100", "sort_column": 2, "order": "asc", "has_header": true}
```
Multi-column sort:
```json
{
  "filename": "data.xlsx",
  "edit_action": "sort",
  "range": "A1:J100",
  "sort_keys": [{"column": 2, "order": "desc"}, {"column": 9, "order": "asc"}],
  "has_header": true
}
```

**`auto_filter`** — Apply auto filter
```json
{"filename": "data.xlsx", "edit_action": "auto_filter", "range": "A1:J100", "field": 2, "criteria": ["空调", "冰箱"]}
```
Filter types:
- Multi-value: `["空调", "冰箱"]`
- Exact match: `"A1"`
- Comparison: `">100"`, `"<50"`
- Wildcard: `"*关键字*"`

**`clear_filter`** — Clear all filters
```json
{"filename": "data.xlsx", "edit_action": "clear_filter", "range": "A1:J100"}
```

**`find_replace`** — Find and replace
```json
{"filename": "data.xlsx", "edit_action": "find_replace", "find": "旧文本", "replace": "新文本", "range": "A1:F20"}
```

**`remove_duplicates`** — Remove duplicate rows
```json
{"filename": "data.xlsx", "edit_action": "remove_duplicates", "range": "A1:F100", "has_header": true, "columns": [1, 2]}
```

**`text_to_columns`** — Split text into columns
```json
{"filename": "data.xlsx", "edit_action": "text_to_columns", "range": "A2:A100", "data_type": "delimited", "comma": true}
```

**`fill_series`** / **`auto_fill`** — Fill series
```json
{"filename": "data.xlsx", "edit_action": "fill_series", "range": "A1:A20", "type": "linear", "step": 1}
{"filename": "data.xlsx", "edit_action": "auto_fill", "source": "A1:A3", "destination": "A1:A20"}
```

**`subtotal`** — Subtotal by group
```json
{"filename": "data.xlsx", "edit_action": "subtotal", "range": "A1:F50", "group_by": 1, "function": "sum", "totals_columns": [4, 5]}
```

**`group`** / **`ungroup`** — Create outline groups
```json
{"filename": "data.xlsx", "edit_action": "group", "rows": "5:10"}
{"filename": "data.xlsx", "edit_action": "ungroup", "rows": "5:10"}
```

**`copy_paste`** — Copy and paste range
```json
{"filename": "data.xlsx", "edit_action": "copy_paste", "source": "A1:C10", "destination": "E1"}
```

**`auto_fit`** — Auto-fit columns/rows
```json
{"filename": "data.xlsx", "edit_action": "auto_fit", "fit_type": "both", "range": "A:D"}
```

---

#### 4.5 Conditional Formatting

**`conditional_format`** — Apply conditional formatting (cell value based)
```json
{
  "filename": "data.xlsx",
  "edit_action": "conditional_format",
  "range": "N2:N128",
  "cf_type": "cell_value",
  "operator": 3,
  "formula1": "业绩差",
  "format": {"fill_color": "浅红", "font_color": "深红"}
}
```

Operators: `1`=between, `2`=not_between, `3`=equal, `4`=not_equal, `5`=greater, `6`=less, `7`=greater_equal, `8`=less_equal

**`conditional_format_v2`** — Extended conditional formatting
```json
{
  "filename": "data.xlsx",
  "edit_action": "conditional_format_v2",
  "range": "B2:B100",
  "cf_type": "duplicate",
  "format": {"fill_color": "浅黄"}
}
```
Types: `top`, `bottom`, `above_average`, `below_average`, `text_contains`, `text_not_contains`, `text_begins_with`, `text_ends_with`, `blank`, `no_blank`, `error`, `no_error`, `time_period`, `duplicate`, `unique`

**`list_conditional_formats`** — List all conditional formats
```json
{"filename": "data.xlsx", "edit_action": "list_conditional_formats", "range": "A1:F100"}
```

**`delete_conditional_format`** — Delete conditional format rules
```json
{"filename": "data.xlsx", "edit_action": "delete_conditional_format", "range": "N2:N128"}
```

---

#### 4.6 Data Validation

**`data_validation`** — Simple list validation
```json
{"filename": "data.xlsx", "edit_action": "data_validation", "range": "B2:B100", "validation_type": "list", "items": ["男", "女"]}
```

**`data_validation_v2`** — Full data validation
```json
{
  "filename": "data.xlsx",
  "edit_action": "data_validation_v2",
  "range": "C2:C100",
  "validation_type": "whole_number",
  "operator": "between",
  "formula1": "1",
  "formula2": "100",
  "error_title": "超出范围",
  "error_message": "必须 1-100",
  "input_title": "输入提示",
  "input_message": "请输入 1-100 的整数"
}
```
Types: `list`, `whole_number`, `decimal`, `date`, `time`, `text_length`, `custom`, `any`

**`circle_invalid_data`** / **`clear_invalid_circles`**
```json
{"filename": "data.xlsx", "edit_action": "circle_invalid_data"}
```

---

#### 4.7 Charts

**`create_chart`** — Create a chart
```json
{
  "filename": "data.xlsx",
  "edit_action": "create_chart",
  "range": "B2:C4",
  "chart_type": "column",
  "position": "A7:F22",
  "title": "Sales by Quarter",
  "show_data_labels": true
}
```
Types: `column`, `bar`, `line`, `pie`, `scatter`, `area`, `doughnut`, `radar`

**`list_charts`** — List all charts
```json
{"filename": "data.xlsx", "edit_action": "list_charts"}
```

**`edit_chart`** — Edit chart properties (deep editing)
```json
{
  "filename": "data.xlsx",
  "edit_action": "edit_chart",
  "chart_index": 1,
  "chart_type": "line",
  "title": "Trend Analysis",
  "show_legend": false,
  "show_data_labels": true,
  "series_format": [{"index": 1, "line_color": "红", "line_width": 2.5, "smooth": true}],
  "trendlines": [{"series_index": 1, "type": "linear", "show_equation": true}],
  "x_axis": {"title": "Month", "tick_label_rotation": -45},
  "y_axis": {"title": "Revenue", "min": 0, "max": 100000},
  "gridlines": {"major_y": true, "minor_y": false}
}
```

**`delete_chart`** / **`chart_style`** / **`export_chart`**
```json
{"filename": "data.xlsx", "edit_action": "delete_chart", "chart_index": 1}
{"filename": "data.xlsx", "edit_action": "chart_style", "chart_index": 1, "chart_style": 5}
{"filename": "data.xlsx", "edit_action": "export_chart", "chart_index": 1, "path": "D:/chart.png"}
```

---

#### 4.8 Tables (ListObject)

**`convert_to_table`** — Convert range to Excel table
```json
{
  "filename": "data.xlsx",
  "edit_action": "convert_to_table",
  "range": "A1:F50",
  "has_headers": true,
  "name": "SalesTable",
  "style": "TableStyleMedium2"
}
```

**`list_tables`** / **`table_style`** / **`table_total_row`** / **`table_resize`** / **`convert_to_range`**
```json
{"filename": "data.xlsx", "edit_action": "list_tables"}
{"filename": "data.xlsx", "edit_action": "table_style", "table_name": "SalesTable", "style": "light_9", "show_totals": true}
{"filename": "data.xlsx", "edit_action": "table_total_row", "table_name": "SalesTable", "show": true, "totals_calculations": [{"column": "金额", "function": "sum"}]}
{"filename": "data.xlsx", "edit_action": "table_resize", "table_name": "SalesTable", "range": "A1:F100"}
{"filename": "data.xlsx", "edit_action": "convert_to_range", "table_name": "SalesTable"}
```

---

#### 4.9 Shapes & Textboxes

**`insert_shape`** — Insert a shape
```json
{
  "filename": "data.xlsx",
  "edit_action": "insert_shape",
  "cell": "B5",
  "shape_type": "rectangle",
  "width": 100,
  "height": 80,
  "fill_color": "蓝",
  "text": "标注",
  "text_font": {"bold": true, "color": "白"}
}
```
Types: `rectangle`, `oval`, `triangle`, `star_5`, `arrow_right`, `callout_rectangle`, `line`, etc.

**`insert_textbox`** — Insert a textbox
```json
{
  "filename": "data.xlsx",
  "edit_action": "insert_textbox",
  "cell": "A1",
  "width": 200,
  "height": 50,
  "text": "说明文字",
  "font": {"name": "微软雅黑", "size": 11, "bold": true},
  "fill_color": "黄",
  "border_visible": true
}
```

**`list_shapes`** / **`edit_shape`** / **`delete_shape`** / **`replace_image`** / **`move_shape`**
```json
{"filename": "data.xlsx", "edit_action": "list_shapes"}
{"filename": "data.xlsx", "edit_action": "edit_shape", "name": "Rectangle 1", "fill_color": "红"}
{"filename": "data.xlsx", "edit_action": "delete_shape", "name": "Rectangle 1"}
{"filename": "data.xlsx", "edit_action": "replace_image", "name": "Picture 1", "path": "D:/new.png"}
{"filename": "data.xlsx", "edit_action": "move_shape", "name": "Rectangle 1", "left": 200, "top": 100}
```

---

#### 4.10 Page Setup & Print

**`page_setup`** — Configure page settings
```json
{
  "filename": "data.xlsx",
  "edit_action": "page_setup",
  "orientation": "landscape",
  "paper_size": "a4",
  "top_margin": 2.54,
  "bottom_margin": 2.54,
  "left_margin": 3.17,
  "right_margin": 3.17,
  "center_horizontally": true,
  "zoom": 80
}
```

**`print_area`** / **`print_titles`** / **`page_break`** / **`set_header`** / **`set_footer`** / **`print_preview`**
```json
{"filename": "data.xlsx", "edit_action": "print_area", "range": "A1:F100"}
{"filename": "data.xlsx", "edit_action": "print_titles", "rows": "$1:$2"}
{"filename": "data.xlsx", "edit_action": "page_break", "operation": "add", "cell": "A20"}
{"filename": "data.xlsx", "edit_action": "set_header", "center": "&A", "right": "&D"}
{"filename": "data.xlsx", "edit_action": "set_footer", "center": "Page &P of &N"}
```

---

#### 4.11 Sheet Appearance

**`set_tab_color`** / **`set_zoom`** / **`set_gridlines_visible`** / **`set_headings_visible`**
```json
{"filename": "data.xlsx", "edit_action": "set_tab_color", "color": "红"}
{"filename": "data.xlsx", "edit_action": "set_zoom", "zoom": 120}
{"filename": "data.xlsx", "edit_action": "set_gridlines_visible", "visible": false}
```

**`move_sheet`** / **`split`** / **`unsplit`** / **`calculate`** / **`set_calculation_mode`**
```json
{"filename": "data.xlsx", "edit_action": "move_sheet", "before": "Sheet2"}
{"filename": "data.xlsx", "edit_action": "freeze_panes", "cell": "B2"}
{"filename": "data.xlsx", "edit_action": "calculate", "scope": "sheet"}
```

**`hide_rows`** / **`show_rows`** / **`hide_cols`** / **`show_cols`**
```json
{"filename": "data.xlsx", "edit_action": "hide_rows", "rows": "5:10"}
{"filename": "data.xlsx", "edit_action": "show_cols", "cols": "C:E"}
```

---

#### 4.12 Other

**`add_comment`** / **`delete_comment`**
```json
{"filename": "data.xlsx", "edit_action": "add_comment", "cell": "A1", "text": "审核通过"}
```

**`add_hyperlink`**
```json
{"filename": "data.xlsx", "edit_action": "add_hyperlink", "cell": "A1", "url": "https://example.com", "text": "Click here"}
```

**`protect`** / **`unprotect`**
```json
{"filename": "data.xlsx", "edit_action": "protect", "password": "123456"}
```

**`clear_contents`** / **`clear_all`**
```json
{"filename": "data.xlsx", "edit_action": "clear_contents", "range": "A1:F20"}
{"filename": "data.xlsx", "edit_action": "clear_all", "range": "A1:F20"}
```

**`set_workbook_property`**
```json
{"filename": "data.xlsx", "edit_action": "set_workbook_property", "title": "销售报告", "author": "张三"}
```

**`trace_precedents`** / **`trace_dependents`** / **`clear_arrows`**
```json
{"filename": "data.xlsx", "edit_action": "trace_precedents", "cell": "E5"}
```

---

### Tool 5: `manage_sheets`

Manage worksheets in a workbook.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `filename` | string | ✅ | Workbook name or path |
| `sheet_action` | string | ✅ | `add`, `delete`, `rename`, `copy`, `hide`, `show`, `activate` |
| `name` | string | * | Sheet name (required for delete/copy/hide/show/activate) |
| `old_name` | string | * | Old name (required for rename) |
| `new_name` | string | * | New name (required for rename) |
| `after` | string | ❌ | Copy destination (sheet name to copy after) |

**Examples:**
```json
{"filename": "data.xlsx", "sheet_action": "add", "name": "Summary"}
{"filename": "data.xlsx", "sheet_action": "delete", "name": "Temp"}
{"filename": "data.xlsx", "sheet_action": "rename", "old_name": "Sheet1", "new_name": "Sales"}
{"filename": "data.xlsx", "sheet_action": "copy", "name": "Sales", "after": "Summary"}
{"filename": "data.xlsx", "sheet_action": "hide", "name": "Backup"}
{"filename": "data.xlsx", "sheet_action": "activate", "name": "Sales"}
```

---

## 中文参考

### 概述

Excel MCP Server 通过模型上下文协议（MCP）暴露 5 个工具。所有操作针对**当前已打开的** Excel 工作簿（通过 COM API）。

---

### 工具 1：`list_workbooks`

列出所有当前打开的 Excel 工作簿。

**参数：** 无

---

### 工具 2：`list_sheets`

列出工作簿中的所有工作表及其使用范围。

**参数：**

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `filename` | string | ✅ | 工作簿名称、完整路径或部分匹配 |

---

### 工具 3：`read_sheet`

从 Excel 工作表读取数据。支持读取值、样式、图表、表格、形状、超链接、数据验证、条件格式、筛选和打印设置。

**⚠️ 始终先调用 `list_sheets` 获取 `used_range`，然后使用 `range` 参数读取大表的特定区域。**

**参数：**

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `filename` | string | ✅ | - | 工作簿名称或路径 |
| `sheet` | string | ❌ | 活动工作表 | 工作表名称 |
| `range` | string | ❌ | UsedRange | 单元格范围如 `A1:F20` |
| `include_style` | boolean | ❌ | false | 包含字体/颜色/对齐/边框信息 |
| `max_rows` | integer | ❌ | 500/200 | 最大行数（有样式200，无样式500） |
| `max_cells` | integer | ❌ | 8000/2000 | 最大单元格数 |
| `head_rows` | integer | ❌ | - | 前 N 行（用于大表） |
| `tail_rows` | integer | ❌ | - | 后 M 行 |
| `include_charts` | boolean | ❌ | false | 返回图表列表 |
| `include_tables` | boolean | ❌ | false | 返回 ListObject 表格 |
| `include_shapes` | boolean | ❌ | false | 返回形状/文本框 |
| `include_hyperlinks` | boolean | ❌ | false | 返回超链接 |
| `include_data_validations` | boolean | ❌ | false | 返回数据验证规则 |
| `include_conditional_formats` | boolean | ❌ | false | 返回条件格式 |
| `include_filters` | boolean | ❌ | false | 返回自动筛选状态 |
| `include_print_settings` | boolean | ❌ | false | 返回打印设置 |

---

### 工具 4：`edit_sheet`

主要编辑工具，包含 **90+ 种操作**。

**常用参数：**

| 参数 | 类型 | 说明 |
|------|------|------|
| `filename` | string | **必填。** 工作簿名称或路径 |
| `edit_action` | string | **必填。** 操作类型（见下方完整列表） |
| `sheet` | string | 工作表名称（默认活动工作表） |
| `operations` | array | 操作对象数组（用于批量操作） |
| `range` | string | 单元格范围如 `A1:F20` |
| `cell` | string | 单个单元格如 `G3` |
| `style` | object | 样式对象（见样式部分） |

**`edit_action` 完整列表：**

| 类别 | 操作 | 说明 |
|------|------|------|
| **值与公式** | `set_value` | 设置单元格值 |
| | `set_formula` | 设置公式 |
| | `convert_text_numbers` | 文本数字转真数字 |
| **样式格式** | `set_style` | 应用样式 |
| | `number_format` | 设置数字格式 |
| | `apply_cell_style` | 应用内置样式 |
| | `list_cell_styles` | 列出可用样式 |
| | `gradient_fill` | 渐变填充 |
| | `pattern_fill` | 图案填充 |
| | `tint_and_shade` | 调整明暗 |
| | `diagonal_borders` | 斜线表头 |
| | `set_row_height` | 设置行高 |
| | `set_column_width` | 设置列宽 |
| | `copy_format` | 格式刷 |
| **布局** | `merge` / `unmerge` | 合并/取消合并 |
| | `insert_rows` / `delete_rows` | 插入/删除行 |
| | `insert_cols` / `delete_cols` | 插入/删除列 |
| **数据操作** | `sort` | 排序 |
| | `auto_filter` / `clear_filter` | 自动筛选/清除筛选 |
| | `find_replace` | 查找替换 |
| | `remove_duplicates` | 删除重复值 |
| | `text_to_columns` | 分列 |
| | `fill_series` / `auto_fill` | 填充序列 |
| | `subtotal` / `remove_subtotal` | 分类汇总 |
| | `group` / `ungroup` | 分组/取消分组 |
| | `copy_paste` | 复制粘贴 |
| | `auto_fit` | 自动调整列宽/行高 |
| **条件格式** | `conditional_format` | 条件格式（基础） |
| | `conditional_format_v2` | 条件格式（扩展） |
| | `list_conditional_formats` | 列出条件格式 |
| | `delete_conditional_format` | 删除条件格式 |
| **数据验证** | `data_validation` | 数据验证（序列） |
| | `data_validation_v2` | 数据验证（完整） |
| | `circle_invalid_data` | 圈出无效数据 |
| | `clear_invalid_circles` | 清除无效数据圈 |
| **图表** | `create_chart` | 创建图表 |
| | `list_charts` | 列出图表 |
| | `edit_chart` | 编辑图表（深度） |
| | `delete_chart` | 删除图表 |
| | `chart_style` | 应用图表样式 |
| | `export_chart` | 导出图表为图片 |
| **表格** | `list_tables` | 列出表格 |
| | `convert_to_table` | 转为表格 |
| | `table_style` | 表格样式 |
| | `table_total_row` | 汇总行 |
| | `table_resize` | 调整表格大小 |
| | `convert_to_range` | 转为普通区域 |
| **形状** | `list_shapes` | 列出形状 |
| | `insert_shape` | 插入形状 |
| | `insert_textbox` | 插入文本框 |
| | `edit_shape` | 编辑形状 |
| | `delete_shape` | 删除形状 |
| | `replace_image` | 替换图片 |
| | `move_shape` | 移动形状 |
| **页面打印** | `page_setup` | 页面设置 |
| | `print_area` | 打印区域 |
| | `print_titles` | 顶端标题 |
| | `page_break` | 分页符 |
| | `set_header` / `set_footer` | 页眉/页脚 |
| | `print_preview` | 打印预览 |
| **工作表外观** | `set_tab_color` | 标签颜色 |
| | `set_zoom` | 缩放 |
| | `set_gridlines_visible` | 网格线 |
| | `set_headings_visible` | 行列标题 |
| | `move_sheet` | 移动工作表 |
| | `split` / `unsplit` | 拆分窗口 |
| | `calculate` | 强制计算 |
| | `set_calculation_mode` | 计算模式 |
| | `freeze_panes` / `unfreeze_panes` | 冻结窗格 |
| | `hide_rows` / `show_rows` | 隐藏/显示行 |
| | `hide_cols` / `show_cols` | 隐藏/显示列 |
| **其他** | `add_comment` / `delete_comment` | 批注 |
| | `add_hyperlink` | 超链接 |
| | `protect` / `unprotect` | 保护 |
| | `clear_contents` / `clear_all` | 清除内容 |
| | `set_workbook_property` | 文档属性 |
| | `trace_precedents` / `trace_dependents` | 公式审核 |
| | `clear_arrows` | 清除追踪箭头 |
| | `insert_image` | 插入图片 |
| | `named_range` | 命名范围 |

---

### 工具 5：`manage_sheets`

管理工作表。

**参数：**

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `filename` | string | ✅ | 工作簿名称或路径 |
| `sheet_action` | string | ✅ | `add`、`delete`、`rename`、`copy`、`hide`、`show`、`activate` |
| `name` | string | * | 工作表名称 |
| `old_name` | string | * | 旧名称（rename 时必填） |
| `new_name` | string | * | 新名称（rename 时必填） |
| `after` | string | ❌ | 复制目标位置 |

**示例：**
```json
{"filename": "data.xlsx", "sheet_action": "add", "name": "汇总"}
{"filename": "data.xlsx", "sheet_action": "delete", "name": "临时"}
{"filename": "data.xlsx", "sheet_action": "rename", "old_name": "Sheet1", "new_name": "销售"}
{"filename": "data.xlsx", "sheet_action": "copy", "name": "销售", "after": "汇总"}
{"filename": "data.xlsx", "sheet_action": "hide", "name": "备份"}
{"filename": "data.xlsx", "sheet_action": "activate", "name": "销售"}
```
