# Excel MCP Server

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.10+-green.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)]()

> 🇨🇳 中文 | 🇺🇸 [English](README.md)

一个用于 Microsoft Excel 自动化的 **MCP (模型上下文协议) 服务器**。让 AI 助手通过自然语言直接操控 Excel —— 读取单元格、写入公式、应用格式、创建图表、管理工作表，以及 **90+ 种操作**。

## 功能特性

### 读写操作
- 读取单元格值、公式、样式、合并单元格
- 写入值和公式，自动类型检测
- 复制、粘贴、查找替换

### 格式与样式
- 字体、颜色、对齐、边框、数字格式
- 条件格式（单元格值、公式、数据条、色阶、图标集、前/后N项、重复值等）
- 单元格样式（渐变填充、图案填充、明暗调整、斜线表头、内置样式）
- 行高、列宽、格式刷

### 数据操作
- 排序（单列/多列）、自动筛选
- 删除重复值、分列、填充序列
- 分类汇总、分组/取消分组、分级显示
- 数据验证（序列、整数、小数、日期、时间、文本长度、自定义公式）

### 图表
- 创建、编辑、删除图表（柱状、条形、折线、饼图、散点、面积、环形、雷达）
- 深度编辑：趋势线、单点变色、数据标签、坐标轴配置、网格线
- 导出图表为图片

### 表格 (ListObject)
- 创建、设置样式、调整大小
- 汇总行（自定义聚合函数）
- 表格与普通区域互转

### 形状与文本框
- 插入、编辑、移动、删除形状（矩形、椭圆、三角形、五角星、箭头等）
- 插入和编辑文本框，支持完整格式化

### 页面布局与打印
- 页面设置（方向、纸张大小、边距、缩放、适应页面）
- 打印区域、顶端标题、分页符
- 页眉页脚（支持日期、页码、工作表名等动态字段）

### 工作表管理
- 添加、删除、重命名、复制、隐藏、显示、激活工作表
- 标签颜色、缩放、网格线、行列标题
- 拆分窗口、冻结窗格、计算模式

### 工作簿属性
- 设置文档属性（标题、作者、关键词等）
- 公式审核（追踪引用单元格、追踪从属单元格）
- 命名范围、批注、超链接、保护

## 快速开始

### 前置条件

- **Windows** 系统，已安装并运行 **Microsoft Excel**
- **Python 3.10+**
- 支持 MCP 的 AI 客户端（Claude Desktop、Cursor 等）

### 安装

```bash
git clone https://github.com/shuncongci/excel-mcp-server.git
cd excel-mcp-server
pip install -e .
```

### 配置 MCP 客户端

**Claude Code**：

```bash
claude mcp add excel-mcp-server -s user -- uv run --directory C:/path/to/excel-mcp-server excel-mcp-server
```

也可以使用等价的 JSON 配置：

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

**Claude Desktop** (`claude_desktop_config.json`)：
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

**Cursor** (`.cursor/mcp.json`)：
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

## 工具列表

| 工具 | 说明 |
|------|------|
| `list_workbooks` | 列出所有打开的 Excel 工作簿 |
| `list_sheets` | 列出工作表及使用范围 |
| `read_sheet` | 读取数据、样式、图表、表格、形状、验证规则、筛选、打印设置 |
| `edit_sheet` | 90+ 种编辑操作（值、公式、格式、图表、表格、形状、页面布局等） |
| `manage_sheets` | 添加、删除、重命名、复制、隐藏、显示、激活工作表 |

完整 API 参考请见 [docs/TOOLS.md](docs/TOOLS.md)。

## 技术架构

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

## 项目结构

```
excel-mcp-server/
├── .gitignore              # Git 忽略规则
├── README.md              # English documentation
├── README_CN.md           # 中文文档
├── LICENSE                # Apache 2.0
├── pyproject.toml         # Python packaging config
├── docs/
│   └── TOOLS.md           # 完整 API 参考（中英文）
├── tests/                 # 回归测试与元数据测试
└── src/
    └── excel_mcp/
        ├── __init__.py    # 包初始化
        ├── server.py      # MCP 服务器（5 个工具）
        └── excel_core.py  # Excel 操作引擎（90+ 操作）
```

## 开发与测试

运行回归测试：

```bash
pip install -e ".[dev]"
pytest
```

测试覆盖模块导入、MCP 工具 schema、License 元数据一致性，以及防止 Python 代码中再次误写小写 JSON 常量（`true` / `false` / `null`）。

## 协议

Apache 2.0 — 详见 [LICENSE](LICENSE)

Copyright 2026 shuncongci
