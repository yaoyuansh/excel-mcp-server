#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Core - Excel 核心操作引擎
通过 win32com COM API 操作当前打开的 Excel 文档
支持：读取、编辑、格式化、条件格式、图表、表格、形状、页面打印等 90+ 操作

用法：
  from excel_core import excel_read, excel_edit, excel_list_sheets, excel_manage_sheets

所有输入输出均为 JSON 格式
"""

import sys
import json
import os
import re
import traceback
import io
import base64
import mimetypes
import tempfile
import zipfile
import uuid
import urllib.parse
import urllib.request

# 设置 UTF-8 输出
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')

import pythoncom
import win32com.client as win32
from decimal import Decimal
from PIL import Image


MAX_INLINE_IMAGE_BASE64_CHARS = 10 * 1024 * 1024 - 4096
IMAGE_SCALE_STEPS = (1.0, 0.92, 0.85, 0.75, 0.66, 0.58, 0.5, 0.42, 0.35, 0.28)
JPEG_QUALITY_STEPS = (92, 86, 80, 74, 68, 62, 56, 50, 44, 38, 32, 26)
PNG_COLOR_STEPS = (None, 256, 128, 64, 32)
REMOTE_IMAGE_DOWNLOAD_TIMEOUT_SECONDS = 60
REMOTE_IMAGE_CONTENT_TYPE_EXTENSION_MAP = {
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
    "image/webp": ".webp",
    "image/svg+xml": ".svg",
    "image/tiff": ".tiff",
    "image/x-icon": ".ico",
    "image/heic": ".heic",
    "image/heif": ".heif",
}

# =====================================================
# Excel 常量映射
# =====================================================
HALIGN_MAP = {
    "left": -4131,      # xlLeft
    "center": -4108,    # xlCenter
    "centre": -4108,    # 英式拼法
    "middle": -4108,    # 别名
    "right": -4152,     # xlRight
    "justify": -4130,   # xlJustify
    "general": 1,       # xlGeneral
    "居中": -4108,
    "左对齐": -4131, "左": -4131,
    "右对齐": -4152, "右": -4152,
    "两端对齐": -4130,
    "常规": 1,
}
VALIGN_MAP = {
    "top": -4160,       # xlTop
    "center": -4108,    # xlCenter
    "centre": -4108,    # 英式拼法
    "middle": -4108,    # 别名
    "bottom": -4107,    # xlBottom
    "顶端": -4160, "顶部": -4160, "上": -4160,
    "居中": -4108, "中部": -4108, "中": -4108,
    "底端": -4107, "底部": -4107, "下": -4107,
}
BORDER_INDEX_MAP = {
    "left": 7,          # xlEdgeLeft
    "right": 10,        # xlEdgeRight
    "top": 8,           # xlEdgeTop
    "bottom": 9,        # xlEdgeBottom
    "inside_h": 12,     # xlInsideHorizontal
    "inside_v": 11,     # xlInsideVertical
    "diagonal_down": 5, # xlDiagonalDown
    "diagonal_up": 6,   # xlDiagonalUp
}
BORDER_STYLE_MAP = {
    "continuous": 1,     # xlContinuous
    "solid": 1,          # 别名：实线
    "single": 1,         # 别名：单实线（Excel 常用）
    "dash": -4115,       # xlDash
    "dashed": -4115,     # 别名
    "dot": -4118,        # xlDot
    "dotted": -4118,     # 别名
    "double": -4119,     # xlDouble
    "thick": 4,          # xlThick（粗线）
    "thin": 2,           # xlThin（细线）
    "medium": 1,
    "none": -4142,       # xlLineStyleNone
}
BORDER_WEIGHT_MAP = {
    "hairline": 1,       # xlHairline
    "extra_thin": 1,     # 别名
    "thin": 2,           # xlThin
    "normal": 2,         # 别名
    "medium": -4138,     # xlMedium
    "thick": 4,          # xlThick
    "bold": 4,           # 别名
    "发丝": 1,
    "细": 2, "细线": 2,
    "中": -4138, "中等": -4138, "中线": -4138,
    "粗": 4, "粗线": 4,
}
BORDER_EDGE_READ_MAP = {name: BORDER_INDEX_MAP[name] for name in ("top", "right", "bottom", "left")}
BORDER_ALL_EDGE_INDEXES = tuple(BORDER_INDEX_MAP[name] for name in ("left", "top", "bottom", "right", "inside_v", "inside_h"))
EXCEL_BORDER_STYLE_NAME_MAP = {
    -4142: "none",
    0: "none",
    1: "continuous",
    -4115: "dash",
    -4118: "dot",
    -4119: "double",
    4: "dash_dot",
    5: "dash_dot_dot",
    13: "slant_dash_dot",
}
EXCEL_BORDER_WEIGHT_NAME_MAP = {
    1: "hairline",
    2: "thin",
    -4138: "medium",
    4: "thick",
}
UNDERLINE_MAP = {
    "single": 2,         # xlUnderlineStyleSingle
    "double": -4119,     # xlUnderlineStyleDouble
    "single_accounting": 4,   # xlUnderlineStyleSingleAccounting
    "double_accounting": 5,   # xlUnderlineStyleDoubleAccounting
    "none": -4142,       # xlUnderlineStyleNone
    "单": 2, "单下划线": 2,
    "双": -4119, "双下划线": -4119,
    "无": -4142,
}
COLOR_NAME_MAP = {
    "红": (255, 0, 0), "red": (255, 0, 0),
    "绿": (0, 128, 0), "green": (0, 128, 0),
    "蓝": (0, 0, 255), "blue": (0, 0, 255),
    "黄": (255, 255, 0), "yellow": (255, 255, 0),
    "黑": (0, 0, 0), "black": (0, 0, 0),
    "白": (255, 255, 255), "white": (255, 255, 255),
    "灰": (128, 128, 128), "gray": (128, 128, 128),
    "橙": (255, 165, 0), "orange": (255, 165, 0),
    "紫": (128, 0, 128), "purple": (128, 0, 128),
    "粉": (255, 192, 203), "pink": (255, 192, 203),
    "棕": (139, 69, 19), "brown": (139, 69, 19),
    "青": (0, 255, 255), "cyan": (0, 255, 255),
    # Excel 条件格式预设 6 色（"突出显示单元格规则"对话框中的标准色）
    "浅红": (255, 199, 206), "浅红色": (255, 199, 206),
    "浅红填充": (255, 199, 206), "浅红色填充": (255, 199, 206),
    "light_red": (255, 199, 206), "light red": (255, 199, 206),
    "深红": (156, 0, 6), "深红色": (156, 0, 6), "dark_red": (156, 0, 6),
    "浅黄": (255, 235, 156), "浅黄色": (255, 235, 156),
    "浅黄填充": (255, 235, 156), "浅黄色填充": (255, 235, 156),
    "light_yellow": (255, 235, 156), "light yellow": (255, 235, 156),
    "深黄": (156, 101, 0), "深黄色": (156, 101, 0),
    "浅绿": (198, 239, 206), "浅绿色": (198, 239, 206),
    "浅绿填充": (198, 239, 206), "浅绿色填充": (198, 239, 206),
    "light_green": (198, 239, 206), "light green": (198, 239, 206),
    "深绿": (0, 97, 0), "深绿色": (0, 97, 0), "dark_green": (0, 97, 0),
    "浅蓝": (221, 235, 247), "浅蓝色": (221, 235, 247), "light_blue": (221, 235, 247), "light blue": (221, 235, 247),
    "深蓝": (0, 51, 102), "深蓝色": (0, 51, 102), "dark_blue": (0, 51, 102), "dark blue": (0, 51, 102),
    "天蓝": (91, 155, 213), "天蓝色": (91, 155, 213), "sky_blue": (91, 155, 213), "sky blue": (91, 155, 213),
    "浅灰": (217, 217, 217), "浅灰色": (217, 217, 217), "light_gray": (217, 217, 217), "light grey": (217, 217, 217),
    "深灰": (89, 89, 89), "深灰色": (89, 89, 89), "dark_gray": (89, 89, 89), "dark grey": (89, 89, 89),
    "金色": (255, 192, 0), "金": (255, 192, 0), "gold": (255, 192, 0),
    "玫红": (255, 0, 102), "玫红色": (255, 0, 102), "magenta": (255, 0, 255),
}

# Word 常量
WORD_HALIGN_MAP = {
    "left": 0,           # wdAlignParagraphLeft
    "center": 1,         # wdAlignParagraphCenter
    "centre": 1,         # 英式拼法
    "middle": 1,         # 别名
    "right": 2,          # wdAlignParagraphRight
    "justify": 3,        # wdAlignParagraphJustify
    "居中": 1,
    "左对齐": 0, "左": 0,
    "右对齐": 2, "右": 2,
    "两端对齐": 3,
}
WORD_UNDERLINE_MAP = {
    "single": 1,         # wdUnderlineSingle
    "double": 3,         # wdUnderlineDouble
    "none": 0,           # wdUnderlineNone
    "dotted": 4,         # wdUnderlineDotted
    "thick": 6,          # wdUnderlineThick
    "wave": 11,          # wdUnderlineWavy
    "单": 1, "单下划线": 1,
    "双": 3, "双下划线": 3,
    "无": 0,
    "点": 4, "点线": 4,
    "粗": 6, "粗线": 6,
    "波浪": 11, "波浪线": 11,
}

# 中文字号名 → 磅值映射
CN_FONT_SIZE_MAP = {
    "初号": 42, "小初": 36,
    "一号": 26, "小一": 24,
    "二号": 22, "小二": 18,
    "三号": 16, "小三": 15,
    "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
    "七号": 5.5, "八号": 5,
}

# Word 边框线型映射
WORD_BORDER_STYLE_MAP = {
    "single": 1,          # wdLineStyleSingle
    "double": 3,          # wdLineStyleDouble (双实线)
    "double_wave": 19,    # wdLineStyleDoubleWavy (双曲线/双波浪线)
    "wave": 18,           # wdLineStyleSingleWavy (单波浪线)
    "thick_thin": 5,      # wdLineStyleThickThinLargeGap
    "thin_thick": 6,      # wdLineStyleThinThickLargeGap
    "dash": 2,            # wdLineStyleDashSmallGap
    "dot": 7,             # wdLineStyleDot
    "none": 0,            # wdLineStyleNone
}

# Word 边框位置映射
WORD_BORDER_INDEX_MAP = {
    "top": -1,            # wdBorderTop
    "left": -2,           # wdBorderLeft
    "bottom": -3,         # wdBorderBottom
    "right": -4,          # wdBorderRight
    "all": "all",         # 四边全部
}

# Word 行距类型映射
WORD_LINE_SPACING_RULE_MAP = {
    "single": 0,          # wdLineSpaceSingle
    "1.5": 1,             # wdLineSpace1pt5
    "double": 2,          # wdLineSpaceDouble
    "at_least": 3,        # wdLineSpaceAtLeast (最小值)
    "exactly": 4,         # wdLineSpaceExactly (固定值)
    "fixed": 4,           # 别名
    "multiple": 5,        # wdLineSpaceMultiple (多倍行距)
}

# =====================================================
# Excel 图表常量（xlChartType）
# =====================================================
CHART_TYPE_MAP = {
    # 柱形 / 条形
    "column": 51, "column_clustered": 51, "簇状柱形图": 51, "柱状图": 51, "柱形图": 51,
    "column_stacked": 52, "堆积柱形图": 52,
    "column_stacked_100": 53, "百分比堆积柱形图": 53,
    "column_3d": -4100, "三维柱形图": -4100,
    "bar": 57, "bar_clustered": 57, "簇状条形图": 57, "条形图": 57,
    "bar_stacked": 58, "堆积条形图": 58,
    "bar_stacked_100": 59, "百分比堆积条形图": 59,
    # 折线
    "line": 4, "折线图": 4,
    "line_stacked": 63, "堆积折线图": 63,
    "line_stacked_100": 64, "百分比堆积折线图": 64,
    "line_markers": 65, "带数据标记折线图": 65,
    "line_markers_stacked": 66, "带数据标记的堆积折线图": 66,
    # 饼图
    "pie": 5, "饼图": 5,
    "pie_3d": -4102, "三维饼图": -4102,
    "pie_of_pie": 68, "复合饼图": 68,
    "bar_of_pie": 71, "复合条饼图": 71,
    "doughnut": -4120, "环形图": -4120,
    "doughnut_exploded": 80, "分离型环形图": 80,
    # 散点 / XY
    "scatter": -4169, "散点图": -4169,
    "scatter_lines": 74, "带直线散点图": 74,
    "scatter_smooth": 72, "带平滑线散点图": 72,
    "scatter_smooth_no_markers": 73, "带平滑线但无标记": 73,
    "scatter_lines_no_markers": 75, "带直线但无标记": 75,
    "bubble": 15, "气泡图": 15,
    "bubble_3d": 87, "三维气泡图": 87,
    # 面积
    "area": 1, "面积图": 1,
    "area_stacked": 76, "堆积面积图": 76,
    "area_stacked_100": 77, "百分比堆积面积图": 77,
    # 雷达
    "radar": -4151, "雷达图": -4151,
    "radar_markers": 81, "带数据标记雷达图": 81,
    "radar_filled": 82, "填充雷达图": 82,
    # 股价/曲面/其它
    "stock_hlc": 88, "盘高盘低盘收图": 88,
    "stock_ohlc": 89, "开盘盘高盘低盘收图": 89,
    "surface": 83, "三维曲面图": 83,
    "combo": -1, "组合图": -1,
}
CHART_TYPE_REVERSE_MAP = {v: k for k, v in CHART_TYPE_MAP.items() if isinstance(v, int) and v != -1}

# 图例位置（xlLegendPosition）
LEGEND_POSITION_MAP = {
    "top": -4160, "顶部": -4160, "上": -4160, "上方": -4160,
    "bottom": -4107, "底部": -4107, "下": -4107, "下方": -4107,
    "left": -4131, "左": -4131, "左侧": -4131,
    "right": -4152, "右": -4152, "右侧": -4152,
    "corner": 2, "右上角": 2,
}
LEGEND_POSITION_REVERSE_MAP = {v: k for k, v in LEGEND_POSITION_MAP.items() if v in {-4160, -4107, -4131, -4152, 2}}

# 数据标签位置（xlDataLabelPosition）
DATA_LABEL_POSITION_MAP = {
    "above": 0, "上方": 0, "顶部": 0,
    "below": 1, "下方": 1, "底部": 1,
    "best_fit": 5, "最佳匹配": 5,
    "center": -4108, "居中": -4108, "中心": -4108,
    "inside_base": 4, "靠内": 4, "数据标签内": 4,
    "inside_end": 3, "轴内侧": 3, "内侧": 3,
    "left": -4131, "左": -4131,
    "right": -4152, "右": -4152,
    "outside_end": 2, "数据标签外": 2, "轴外侧": 2, "外侧": 2,
}

# Excel 内置 CellStyle 名（覆盖最常用的）
EXCEL_BUILTIN_CELL_STYLES = {
    "normal": "Normal", "普通": "Normal", "常规": "Normal",
    "bad": "Bad", "差": "Bad", "坏": "Bad",
    "good": "Good", "好": "Good", "佳": "Good",
    "neutral": "Neutral", "中性": "Neutral", "中": "Neutral",
    "calculation": "Calculation", "计算": "Calculation",
    "check_cell": "Check Cell", "检查单元格": "Check Cell",
    "explanatory_text": "Explanatory Text", "解释性文本": "Explanatory Text",
    "input": "Input", "输入": "Input",
    "linked_cell": "Linked Cell", "链接单元格": "Linked Cell",
    "note": "Note", "注释": "Note",
    "output": "Output", "输出": "Output",
    "warning_text": "Warning Text", "警告文本": "Warning Text",
    "title": "Title", "标题": "Title",
    "heading_1": "Heading 1", "标题1": "Heading 1", "标题 1": "Heading 1",
    "heading_2": "Heading 2", "标题2": "Heading 2", "标题 2": "Heading 2",
    "heading_3": "Heading 3", "标题3": "Heading 3", "标题 3": "Heading 3",
    "heading_4": "Heading 4", "标题4": "Heading 4", "标题 4": "Heading 4",
    "total": "Total", "汇总": "Total",
    "accent_1": "Accent1", "强调1": "Accent1", "着色1": "Accent1",
    "accent_2": "Accent2", "强调2": "Accent2", "着色2": "Accent2",
    "accent_3": "Accent3", "强调3": "Accent3", "着色3": "Accent3",
    "accent_4": "Accent4", "强调4": "Accent4", "着色4": "Accent4",
    "accent_5": "Accent5", "强调5": "Accent5", "着色5": "Accent5",
    "accent_6": "Accent6", "强调6": "Accent6", "着色6": "Accent6",
    "20%_accent_1": "20% - Accent1", "20%_accent_2": "20% - Accent2",
    "20%_accent_3": "20% - Accent3", "20%_accent_4": "20% - Accent4",
    "20%_accent_5": "20% - Accent5", "20%_accent_6": "20% - Accent6",
    "40%_accent_1": "40% - Accent1", "40%_accent_2": "40% - Accent2",
    "40%_accent_3": "40% - Accent3", "40%_accent_4": "40% - Accent4",
    "40%_accent_5": "40% - Accent5", "40%_accent_6": "40% - Accent6",
    "60%_accent_1": "60% - Accent1", "60%_accent_2": "60% - Accent2",
    "60%_accent_3": "60% - Accent3", "60%_accent_4": "60% - Accent4",
    "60%_accent_5": "60% - Accent5", "60%_accent_6": "60% - Accent6",
    "currency": "Currency", "货币": "Currency",
    "currency_0": "Currency [0]", "货币[0]": "Currency [0]",
    "comma": "Comma", "千位分隔": "Comma",
    "comma_0": "Comma [0]", "千位分隔[0]": "Comma [0]",
    "percent": "Percent", "百分比": "Percent",
}

# 填充图案（Interior.Pattern, xlPattern）
PATTERN_FILL_MAP = {
    "solid": 1, "实心": 1, "实色": 1,
    "automatic": -4105, "自动": -4105,
    "none": -4142, "无": -4142,
    "gray_75": -4126, "75%灰": -4126, "深色": -4126,
    "gray_50": -4125, "50%灰": -4125, "半色": -4125,
    "gray_25": -4124, "25%灰": -4124, "浅色": -4124,
    "gray_16": 17, "12.5%灰": 17,
    "gray_8": 18, "6.25%灰": 18,
    "horizontal": -4128, "水平线": -4128,
    "vertical": -4166, "垂直线": -4166, "竖线": -4166,
    "diagonal_down": -4121, "对角向下": -4121, "右斜线": -4121,
    "diagonal_up": -4162, "对角向上": -4162, "左斜线": -4162,
    "checker": 9, "方格": 9, "棋盘": 9,
    "semi_gray_75": 10, "深色棋盘": 10,
    "light_horizontal": 11, "细水平线": 11,
    "light_vertical": 12, "细垂直线": 12,
    "light_down": 13, "细下斜线": 13,
    "light_up": 14, "细上斜线": 14,
    "grid": 15, "网格": 15,
    "crisscross": 16, "交叉线": 16,
}

# 渐变方向（xlGradientStyle）
GRADIENT_STYLE_MAP = {
    "horizontal": 1, "水平": 1,
    "vertical": 2, "垂直": 2,
    "diagonal_down": 3, "对角向下": 3,
    "diagonal_up": 4, "对角向上": 4,
    "from_corner": 5, "从角部": 5,
    "from_center": 6, "从中心": 6, "径向": 6,
}

# 页面方向
PAGE_ORIENTATION_MAP = {
    "portrait": 1, "纵向": 1, "竖向": 1, "竖": 1,
    "landscape": 2, "横向": 2, "横": 2,
}

# 纸张大小（xlPaperSize 部分常用）
PAPER_SIZE_MAP = {
    "letter": 1, "信纸": 1,
    "tabloid": 3, "tabloid_extra": 3,
    "ledger": 4,
    "legal": 5, "法律": 5,
    "statement": 6,
    "executive": 7, "executive_paper": 7,
    "a3": 8, "A3": 8,
    "a4": 9, "A4": 9,
    "a4_small": 10,
    "a5": 11, "A5": 11,
    "b4": 12, "B4": 12,
    "b5": 13, "B5": 13,
    "folio": 14,
    "quarto": 15,
    "envelope_10": 20, "信封10号": 20,
    "envelope_dl": 27, "信封DL": 27,
    "envelope_c5": 28,
}

# 填充序列类型（xlAutoFillType 与 xlFillType / xlDataSeriesType）
SERIES_TYPE_MAP = {
    "linear": 0, "等差": 0, "等差数列": 0,
    "growth": 1, "等比": 1, "等比数列": 1,
    "date": 2, "日期": 2,
    "auto_fill": 3, "自动填充": 3,
}
SERIES_DATE_UNIT_MAP = {
    "day": 0, "天": 0, "日": 0,
    "weekday": 1, "工作日": 1,
    "month": 2, "月": 2,
    "year": 3, "年": 3,
}

# 文本分列（xlTextParsingType, xlTextQualifier）
TEXT_PARSING_TYPE_MAP = {
    "delimited": 1, "分隔符号": 1, "分隔符": 1,
    "fixed_width": 2, "固定宽度": 2,
}

# AutoShape 类型（msoAutoShapeType 常用）
AUTOSHAPE_TYPE_MAP = {
    "rectangle": 1, "矩形": 1,
    "rounded_rectangle": 5, "圆角矩形": 5,
    "oval": 9, "椭圆": 9, "圆": 9,
    "diamond": 4, "菱形": 4,
    "triangle": 7, "三角形": 7,
    "right_triangle": 8, "直角三角形": 8,
    "parallelogram": 2, "平行四边形": 2,
    "trapezoid": 3, "梯形": 3,
    "pentagon": 51, "五边形": 51,
    "hexagon": 10, "六边形": 10,
    "octagon": 11, "八边形": 11,
    "star_5": 92, "五角星": 92,
    "arrow_right": 13, "右箭头": 13,
    "arrow_left": 66, "左箭头": 66,
    "arrow_up": 67, "上箭头": 67,
    "arrow_down": 68, "下箭头": 68,
    "callout_rectangle": 105, "矩形标注": 105,
    "callout_oval": 107, "椭圆标注": 107,
    "callout_cloud": 108, "云形标注": 108,
    "heart": 21, "心形": 21,
    "lightning": 22, "闪电": 22,
    "sun": 23, "太阳": 23,
    "moon": 24, "月亮": 24,
    "cloud": 179, "云": 179,
    "line": 9999,  # 特殊值，用 AddLine 而非 AddShape
}


def _resolve_font_size(size_val):
    """将字号值解析为磅值（支持中文字号名、数字、字符串数字）"""
    if isinstance(size_val, (int, float)):
        return size_val
    if isinstance(size_val, str):
        s = size_val.strip()
        # 中文字号名
        if s in CN_FONT_SIZE_MAP:
            return CN_FONT_SIZE_MAP[s]
        # 尝试解析为数字
        try:
            return float(s)
        except ValueError:
            pass
    return None


# Excel 错误码映射（COM 返回的整数 → 可读错误名）
EXCEL_ERROR_MAP = {
    -2146826281: "#DIV/0!",
    -2146826246: "#N/A",
    -2146826259: "#NAME?",
    -2146826288: "#NULL!",
    -2146826252: "#NUM!",
    -2146826265: "#REF!",
    -2146826273: "#VALUE!",
    -2146826245: "#GETTING_DATA",
}


def parse_color(color_input):
    """解析颜色输入 → (R, G, B) 元组
    支持格式：
    - [R, G, B] 或 (R, G, B)
    - 中文颜色名（红/绿/蓝等）或英文名（red/green/blue等）
    - "R,G,B" 或 "R, G, B"
    - "#RRGGBB" 或 "#RGB"
    - "RRGGBB" 裸 hex（6 位）或 "RGB"（3 位）
    - "0xRRGGBB" 前缀
    - "rgb(R,G,B)" / "rgba(R,G,B,A)"（A 会被忽略）
    - 空格/制表符会被去掉
    """
    if isinstance(color_input, (list, tuple)) and len(color_input) == 3:
        return tuple(int(x) for x in color_input)
    if isinstance(color_input, str):
        c = color_input.strip().lower()
        # 先尝试颜色名（注意 COLOR_NAME_MAP 的键含中文，不能 lower 破坏）
        if color_input.strip() in COLOR_NAME_MAP:
            return COLOR_NAME_MAP[color_input.strip()]
        if c in COLOR_NAME_MAP:
            return COLOR_NAME_MAP[c]
        # "rgb(R,G,B)" / "rgba(R,G,B,A)"
        if c.startswith("rgb"):
            m = re.match(r"rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", c)
            if m:
                return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
        # "0xRRGGBB" 前缀
        if c.startswith("0x"):
            c = c[2:]
        # "#RRGGBB" 前缀
        if c.startswith("#"):
            c = c[1:]
        # 去掉中间空格（FF FF FF → FFFFFF），并将中文全角逗号 / 顿号 归一为英文逗号
        c_nospace = c.replace(" ", "").replace("\t", "").replace("，", ",").replace("、", ",")
        # "R,G,B" 格式
        if "," in c_nospace:
            try:
                parts = [int(x.strip()) for x in c_nospace.split(",")]
                if len(parts) >= 3:
                    return tuple(parts[:3])
            except ValueError:
                pass
        # "RRGGBB" 裸 hex（6 位）
        if len(c_nospace) == 6 and all(ch in "0123456789abcdef" for ch in c_nospace):
            return (int(c_nospace[0:2], 16), int(c_nospace[2:4], 16), int(c_nospace[4:6], 16))
        # "RGB" 裸 hex（3 位缩写）
        if len(c_nospace) == 3 and all(ch in "0123456789abcdef" for ch in c_nospace):
            return (int(c_nospace[0]*2, 16), int(c_nospace[1]*2, 16), int(c_nospace[2]*2, 16))
    raise ValueError(f"无法解析颜色: {color_input}，支持格式：颜色名(红/red)、#RRGGBB、RRGGBB、rgb(R,G,B)、R,G,B")


def _resolve_excel_border_style(style_val):
    """归一化 Excel 边框样式值 → 整数常量
    支持：字符串别名（从 BORDER_STYLE_MAP 查）、整数直接返回、数字字符串转整数
    """
    if isinstance(style_val, (int, float)):
        return int(style_val)
    if isinstance(style_val, str):
        s = style_val.strip().lower()
        # 先查映射表
        if s in BORDER_STYLE_MAP:
            return BORDER_STYLE_MAP[s]
        # 尝试作为数字解析
        try:
            return int(s)
        except ValueError:
            pass
    raise ValueError(f"无法解析边框样式: {style_val}，支持的别名: {list(BORDER_STYLE_MAP.keys())}")


def _resolve_from_map(value, mapping, name):
    """通用：归一化字符串/数字到整数映射值。未命中时抛错附合法别名。"""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(value)
    if isinstance(value, str):
        s = value.strip().lower()
        # 先查原始（保留大小写/中文）
        if value in mapping:
            return mapping[value]
        # 再查小写
        if s in mapping:
            return mapping[s]
        # 作为数字解析
        try:
            return int(s)
        except ValueError:
            pass
    raise ValueError(f"无法解析{name}: {value}，支持的别名: {list(mapping.keys())}")


def _resolve_excel_border_weight(weight_val):
    """归一化 Excel 边框粗细 → 整数常量"""
    return _resolve_from_map(weight_val, BORDER_WEIGHT_MAP, "Excel 边框粗细")


def _resolve_excel_underline(underline_val):
    """归一化 Excel 下划线 → 整数常量，支持布尔
    True → single(2), False → none(-4142)
    """
    if isinstance(underline_val, bool):
        return 2 if underline_val else -4142
    return _resolve_from_map(underline_val, UNDERLINE_MAP, "Excel 下划线")


def _resolve_excel_halign(halign_val):
    """归一化 Excel 水平对齐 → 整数常量"""
    return _resolve_from_map(halign_val, HALIGN_MAP, "Excel 水平对齐")


def _resolve_excel_valign(valign_val):
    """归一化 Excel 垂直对齐 → 整数常量"""
    return _resolve_from_map(valign_val, VALIGN_MAP, "Excel 垂直对齐")


def _resolve_word_halign(halign_val):
    """归一化 Word 段落水平对齐 → 整数常量"""
    return _resolve_from_map(halign_val, WORD_HALIGN_MAP, "Word 水平对齐")


def _resolve_word_underline(underline_val):
    """归一化 Word 下划线 → 整数常量，支持布尔
    True → single(1), False → none(0)
    """
    if isinstance(underline_val, bool):
        return 1 if underline_val else 0
    return _resolve_from_map(underline_val, WORD_UNDERLINE_MAP, "Word 下划线")


def _to_bool(value, default=None):
    """规范化布尔值：兼容 True/False/None/"true"/"false"/"是"/"否"/1/0/"1"/"0"
    None 时返回 default（未提供 default 则抛错）
    """
    if isinstance(value, bool):
        return value
    if value is None:
        if default is None:
            raise ValueError("布尔值不能为 None")
        return default
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        s = value.strip().lower()
        if s in ("true", "1", "yes", "y", "是", "开", "启用", "on"):
            return True
        if s in ("false", "0", "no", "n", "否", "关", "禁用", "off", ""):
            return False
    raise ValueError(f"无法解析为布尔值: {value}，支持：true/false/是/否/1/0/yes/no")


def rgb_to_ole(r, g, b):
    """RGB → OLE 颜色值（BGR 格式）"""
    return (b << 16) | (g << 8) | r


def color_to_ole(color_input):
    """颜色输入 → OLE 颜色值"""
    r, g, b = parse_color(color_input)
    return rgb_to_ole(r, g, b)


def ole_to_rgb(ole_color):
    """OLE 颜色值 → (R, G, B)"""
    if ole_color is None or ole_color < 0:
        return None
    ole_color = int(ole_color)
    r = ole_color & 0xFF
    g = (ole_color >> 8) & 0xFF
    b = (ole_color >> 16) & 0xFF
    return (r, g, b)


def _to_int_or_none(value):
    try:
        return int(value)
    except Exception:
        return None


def _read_excel_border(border):
    line_style = _to_int_or_none(border.LineStyle)
    style_name = EXCEL_BORDER_STYLE_NAME_MAP.get(line_style, str(line_style) if line_style is not None else "unknown")
    info = {
        "style": style_name,
        "style_code": line_style,
    }
    if style_name != "none":
        weight_code = _to_int_or_none(border.Weight)
        info["weight"] = EXCEL_BORDER_WEIGHT_NAME_MAP.get(weight_code, str(weight_code) if weight_code is not None else None)
        info["weight_code"] = weight_code
        try:
            info["color"] = ole_to_rgb(border.Color)
        except Exception:
            info["color"] = None
    return info


def _read_excel_cell_borders(cell):
    borders = {}
    for edge_name, edge_idx in BORDER_EDGE_READ_MAP.items():
        try:
            borders[edge_name] = _read_excel_border(cell.Borders(edge_idx))
        except Exception as e:
            borders[edge_name] = {"error": str(e)}
    return borders


def _apply_excel_border_style(border, border_style):
    if not isinstance(border_style, dict):
        return
    if border_style.get("style") is not None:
        border.LineStyle = _resolve_excel_border_style(border_style["style"])
    if border_style.get("weight") is not None:
        border.Weight = _resolve_excel_border_weight(border_style["weight"])
    if border_style.get("color") is not None:
        border.Color = color_to_ole(border_style["color"])


def _resolve_chart_type(value):
    """图表类型字符串/数字 → xlChartType 整数。"""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(value)
    if isinstance(value, str):
        key = value.strip().lower()
        if key in CHART_TYPE_MAP:
            return CHART_TYPE_MAP[key]
        # 大小写敏感的中文键直接匹配
        if value in CHART_TYPE_MAP:
            return CHART_TYPE_MAP[value]
        try:
            return int(value)
        except ValueError:
            pass
    raise ValueError(f"无法解析 chart_type: {value}")


def _chart_type_name(val):
    if val is None:
        return None
    try:
        return CHART_TYPE_REVERSE_MAP.get(int(val), f"type_{val}")
    except Exception:
        return f"type_{val}"


def _apply_chart_axis_config(axis_obj, cfg):
    """统一应用坐标轴配置。
    cfg 支持：title, min, max, major_unit, minor_unit, number_format,
             font:{name,size,bold,italic,color}, line_color, line_visible,
             tick_label_rotation, log_scale, reverse
    """
    if not isinstance(cfg, dict):
        return
    title_text = cfg.get("title")
    if title_text is not None:
        if title_text == "":
            axis_obj.HasTitle = False
        else:
            axis_obj.HasTitle = True
            axis_obj.AxisTitle.Text = str(title_text)
            tf = cfg.get("title_font")
            if isinstance(tf, dict):
                _apply_text_font(axis_obj.AxisTitle.Format.TextFrame2.TextRange.Font, tf)
    if cfg.get("min") is not None:
        try: axis_obj.MinimumScale = float(cfg["min"])
        except Exception: pass
    if cfg.get("max") is not None:
        try: axis_obj.MaximumScale = float(cfg["max"])
        except Exception: pass
    if cfg.get("major_unit") is not None:
        try: axis_obj.MajorUnit = float(cfg["major_unit"])
        except Exception: pass
    if cfg.get("minor_unit") is not None:
        try: axis_obj.MinorUnit = float(cfg["minor_unit"])
        except Exception: pass
    if cfg.get("number_format") is not None:
        try: axis_obj.TickLabels.NumberFormat = str(cfg["number_format"])
        except Exception: pass
    if cfg.get("log_scale") is not None:
        try: axis_obj.ScaleType = 1 if cfg["log_scale"] else 0  # xlScaleLogarithmic=1
        except Exception: pass
    if cfg.get("reverse") is not None:
        try: axis_obj.ReversePlotOrder = bool(cfg["reverse"])
        except Exception: pass
    if cfg.get("tick_label_rotation") is not None:
        try: axis_obj.TickLabels.Orientation = int(cfg["tick_label_rotation"])
        except Exception: pass
    font = cfg.get("font")
    if isinstance(font, dict):
        try:
            _apply_text_font(axis_obj.TickLabels.Format.TextFrame2.TextRange.Font, font)
        except Exception:
            pass
    if cfg.get("line_color") is not None:
        try: axis_obj.Format.Line.ForeColor.RGB = color_to_ole(cfg["line_color"])
        except Exception: pass
    if cfg.get("line_visible") is not None:
        try:
            axis_obj.Format.Line.Visible = -1 if cfg["line_visible"] else 0
        except Exception: pass


def _apply_text_font(font_obj, fcfg):
    """应用 TextFrame2 类型的 Font 配置（用于图表标题/标签/轴标签）"""
    if not isinstance(fcfg, dict):
        return
    if fcfg.get("name"):
        try: font_obj.Name = str(fcfg["name"])
        except Exception: pass
    if fcfg.get("size") is not None:
        try: font_obj.Size = float(fcfg["size"])
        except Exception: pass
    if fcfg.get("bold") is not None:
        try: font_obj.Bold = -1 if fcfg["bold"] else 0
        except Exception: pass
    if fcfg.get("italic") is not None:
        try: font_obj.Italic = -1 if fcfg["italic"] else 0
        except Exception: pass
    if fcfg.get("color") is not None:
        try: font_obj.Fill.ForeColor.RGB = color_to_ole(fcfg["color"])
        except Exception: pass


def _apply_series_format(series_obj, fmt):
    """应用数据系列格式。fmt 支持：
      fill_color, border_color, border_width, line_color, line_width,
      marker_style, marker_size, marker_fill_color, marker_border_color,
      transparency, gap_width, overlap, smooth(line/scatter),
      data_label_format / data_label_position / show_value / show_category_name / show_percentage,
      points: [{index, fill_color, border_color}] 用于单点变色
    """
    if not isinstance(fmt, dict):
        return
    # 填充色
    if fmt.get("fill_color") is not None:
        try:
            series_obj.Format.Fill.Visible = -1
            series_obj.Format.Fill.Solid()
            series_obj.Format.Fill.ForeColor.RGB = color_to_ole(fmt["fill_color"])
        except Exception:
            try:
                series_obj.Interior.Color = color_to_ole(fmt["fill_color"])
            except Exception:
                pass
    # 边框
    if fmt.get("border_color") is not None:
        try:
            series_obj.Format.Line.Visible = -1
            series_obj.Format.Line.ForeColor.RGB = color_to_ole(fmt["border_color"])
        except Exception:
            pass
    if fmt.get("border_width") is not None:
        try: series_obj.Format.Line.Weight = float(fmt["border_width"])
        except Exception: pass
    # 折线
    if fmt.get("line_color") is not None:
        try:
            series_obj.Format.Line.Visible = -1
            series_obj.Format.Line.ForeColor.RGB = color_to_ole(fmt["line_color"])
        except Exception:
            pass
    if fmt.get("line_width") is not None:
        try: series_obj.Format.Line.Weight = float(fmt["line_width"])
        except Exception: pass
    if fmt.get("smooth") is not None:
        try: series_obj.Smooth = bool(fmt["smooth"])
        except Exception: pass
    # 透明度（0-1）
    if fmt.get("transparency") is not None:
        try: series_obj.Format.Fill.Transparency = float(fmt["transparency"])
        except Exception: pass
    # 标记
    if fmt.get("marker_style") is not None:
        marker_map = {
            "none": -4142, "无": -4142,
            "auto": -4105, "自动": -4105,
            "square": 1, "方块": 1,
            "diamond": 2, "菱形": 2,
            "triangle": 3, "三角形": 3,
            "x": 4, "X": 4,
            "star": 5, "星形": 5,
            "dot": 6, "小圆点": 6,
            "dash": 7, "短划线": 7,
            "circle": 8, "圆": 8,
            "plus": 9, "加号": 9,
            "picture": -4147, "图片": -4147,
        }
        try:
            ms = fmt["marker_style"]
            val = marker_map.get(str(ms).lower(), ms)
            series_obj.MarkerStyle = int(val)
        except Exception:
            pass
    if fmt.get("marker_size") is not None:
        try: series_obj.MarkerSize = int(fmt["marker_size"])
        except Exception: pass
    if fmt.get("marker_fill_color") is not None:
        try: series_obj.MarkerBackgroundColor = color_to_ole(fmt["marker_fill_color"])
        except Exception: pass
    if fmt.get("marker_border_color") is not None:
        try: series_obj.MarkerForegroundColor = color_to_ole(fmt["marker_border_color"])
        except Exception: pass
    # 柱形间距 / 重叠（柱形图特有）
    if fmt.get("gap_width") is not None:
        try: series_obj.Parent.GapWidth = int(fmt["gap_width"])
        except Exception: pass
    if fmt.get("overlap") is not None:
        try: series_obj.Parent.Overlap = int(fmt["overlap"])
        except Exception: pass
    # 数据标签
    if "show_value" in fmt or "show_category_name" in fmt or "show_percentage" in fmt or \
       "data_label_format" in fmt or "data_label_position" in fmt:
        try:
            series_obj.HasDataLabels = True
            dl = series_obj.DataLabels()
            if "show_value" in fmt:
                try: dl.ShowValue = bool(fmt["show_value"])
                except Exception: pass
            if "show_category_name" in fmt:
                try: dl.ShowCategoryName = bool(fmt["show_category_name"])
                except Exception: pass
            if "show_percentage" in fmt:
                try: dl.ShowPercentage = bool(fmt["show_percentage"])
                except Exception: pass
            if fmt.get("data_label_format"):
                try: dl.NumberFormat = str(fmt["data_label_format"])
                except Exception: pass
            if fmt.get("data_label_position") is not None:
                pos = fmt["data_label_position"]
                pos_val = DATA_LABEL_POSITION_MAP.get(str(pos).lower(), pos)
                try: dl.Position = int(pos_val)
                except Exception: pass
        except Exception:
            pass
    # 单点变色
    points = fmt.get("points")
    if isinstance(points, list):
        for p in points:
            if not isinstance(p, dict) or "index" not in p:
                continue
            try:
                pt = series_obj.Points(int(p["index"]))
                if p.get("fill_color") is not None:
                    pt.Format.Fill.Visible = -1
                    pt.Format.Fill.Solid()
                    pt.Format.Fill.ForeColor.RGB = color_to_ole(p["fill_color"])
                if p.get("border_color") is not None:
                    pt.Format.Line.Visible = -1
                    pt.Format.Line.ForeColor.RGB = color_to_ole(p["border_color"])
                # 单点的数据标签
                if "show_value" in p or "data_label_format" in p:
                    try:
                        pt.HasDataLabel = True
                        if "show_value" in p:
                            pt.DataLabel.ShowValue = bool(p["show_value"])
                        if p.get("data_label_format"):
                            pt.DataLabel.NumberFormat = str(p["data_label_format"])
                    except Exception:
                        pass
                # 饼图分裂
                if p.get("explosion") is not None:
                    try: pt.Explosion = int(p["explosion"])
                    except Exception: pass
            except Exception:
                pass


def _find_chart_by_identifier(sheet, identifier):
    """通过 1-based 序号、Shape 名（'图表 1'）或图表标题查找图表。
    返回 ShapeObject（含 Chart 属性）或 None。"""
    if identifier is None:
        return None
    shapes = sheet.Shapes
    # 1-based 序号
    if isinstance(identifier, (int, float)) and not isinstance(identifier, bool):
        idx = int(identifier)
        chart_count = 0
        for i in range(1, shapes.Count + 1):
            sh = shapes(i)
            if sh.HasChart:
                chart_count += 1
                if chart_count == idx:
                    return sh
        return None
    name = str(identifier)
    # 先按 Shape Name 精确匹配
    for i in range(1, shapes.Count + 1):
        sh = shapes(i)
        if sh.HasChart and sh.Name == name:
            return sh
    # 再按图表标题匹配
    for i in range(1, shapes.Count + 1):
        sh = shapes(i)
        if sh.HasChart:
            try:
                if sh.Chart.HasTitle and sh.Chart.ChartTitle.Text == name:
                    return sh
            except Exception:
                pass
    return None


def _read_chart_brief(chart_shape, include_series=True):
    """读取一个图表的简要信息（用于 list_charts 与 read_excel 的 include_charts）"""
    info = {
        "shape_name": chart_shape.Name,
        "left": float(chart_shape.Left),
        "top": float(chart_shape.Top),
        "width": float(chart_shape.Width),
        "height": float(chart_shape.Height),
    }
    try:
        chart = chart_shape.Chart
        info["chart_type"] = _chart_type_name(chart.ChartType)
        info["chart_type_code"] = int(chart.ChartType)
    except Exception:
        info["chart_type"] = None
        info["chart_type_code"] = None
    try:
        info["has_title"] = bool(chart.HasTitle)
        info["title"] = chart.ChartTitle.Text if chart.HasTitle else None
    except Exception:
        info["title"] = None
    try:
        info["has_legend"] = bool(chart.HasLegend)
        info["legend_position"] = LEGEND_POSITION_REVERSE_MAP.get(int(chart.Legend.Position)) if chart.HasLegend else None
    except Exception:
        info["legend_position"] = None
    try:
        info["source_data"] = chart.PlotArea.Parent.SeriesCollection(1).Formula
    except Exception:
        info["source_data"] = None
    if include_series:
        series_list = []
        try:
            sc = chart.SeriesCollection()
            for s_idx in range(1, sc.Count + 1):
                s = sc(s_idx)
                s_info = {"index": s_idx}
                try: s_info["name"] = s.Name
                except Exception: pass
                try: s_info["formula"] = s.Formula
                except Exception: pass
                try:
                    s_info["fill_color"] = ole_to_rgb(s.Format.Fill.ForeColor.RGB)
                except Exception:
                    pass
                series_list.append(s_info)
        except Exception:
            pass
        info["series"] = series_list
        info["series_count"] = len(series_list)
    return info


def _excel_to_xy_inches(emu_or_pt, unit="point"):
    """COM Left/Top/Width/Height 默认是 point；EMU 转换备用。"""
    if unit == "emu":
        return emu_or_pt / 914400.0
    return emu_or_pt / 72.0  # point → inch


def _resolve_paper_size(val):
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return int(val)
    if isinstance(val, str):
        key = val.strip().lower()
        if key in PAPER_SIZE_MAP:
            return PAPER_SIZE_MAP[key]
        if val in PAPER_SIZE_MAP:
            return PAPER_SIZE_MAP[val]
        try:
            return int(val)
        except ValueError:
            pass
    return None


def _resolve_orientation(val):
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return int(val)
    if isinstance(val, str):
        key = val.strip().lower()
        if key in PAGE_ORIENTATION_MAP:
            return PAGE_ORIENTATION_MAP[key]
        if val in PAGE_ORIENTATION_MAP:
            return PAGE_ORIENTATION_MAP[val]
    return None


def _resolve_pattern_fill(val):
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return int(val)
    if isinstance(val, str):
        key = val.strip().lower()
        if key in PATTERN_FILL_MAP:
            return PATTERN_FILL_MAP[key]
        if val in PATTERN_FILL_MAP:
            return PATTERN_FILL_MAP[val]
    return None


def _read_excel_list_objects(sheet):
    """读取工作表上的 ListObject（Excel 表格 Ctrl+T）"""
    items = []
    try:
        lo_count = sheet.ListObjects.Count
    except Exception:
        return items
    for i in range(1, lo_count + 1):
        try:
            lo = sheet.ListObjects(i)
            info = {
                "index": i,
                "name": lo.Name,
                "range": lo.Range.Address,
            }
            try: info["style_name"] = lo.TableStyle.Name if lo.TableStyle else None
            except Exception: info["style_name"] = None
            try: info["show_header"] = bool(lo.ShowHeaders)
            except Exception: pass
            try: info["show_totals"] = bool(lo.ShowTotals)
            except Exception: pass
            try: info["show_banded_rows"] = bool(lo.ShowTableStyleRowStripes)
            except Exception: pass
            try: info["show_banded_columns"] = bool(lo.ShowTableStyleColumnStripes)
            except Exception: pass
            items.append(info)
        except Exception:
            pass
    return items


def _read_excel_shapes(sheet, max_shapes=200):
    """读取工作表所有形状/图片/文本框（不含图表，那是单独的 charts 清单）"""
    items = []
    try:
        count = sheet.Shapes.Count
    except Exception:
        return items
    for i in range(1, min(count, max_shapes) + 1):
        try:
            sh = sheet.Shapes(i)
            if sh.HasChart:
                continue
            info = {
                "index": i,
                "name": sh.Name,
                "left": float(sh.Left),
                "top": float(sh.Top),
                "width": float(sh.Width),
                "height": float(sh.Height),
            }
            try:
                info["shape_type"] = int(sh.Type)
            except Exception:
                pass
            try:
                if sh.HasTextFrame:
                    info["text"] = sh.TextFrame2.TextRange.Text
            except Exception:
                pass
            items.append(info)
        except Exception:
            pass
    return items


def _read_excel_hyperlinks(sheet, max_links=500):
    """读取工作表所有超链接"""
    items = []
    try:
        count = sheet.Hyperlinks.Count
    except Exception:
        return items
    for i in range(1, min(count, max_links) + 1):
        try:
            hl = sheet.Hyperlinks(i)
            items.append({
                "cell": hl.Range.Address,
                "address": hl.Address or "",
                "sub_address": hl.SubAddress or "",
                "display": hl.TextToDisplay or "",
                "tooltip": hl.ScreenTip or "",
            })
        except Exception:
            pass
    return items


def _read_excel_data_validations(sheet, sample_range=None):
    """扫描工作表 Validation 规则。需要逐 cell 查询，开销大；默认扫 UsedRange。"""
    items = []
    seen = set()
    try:
        rng = sheet.Range(sample_range) if sample_range else sheet.UsedRange
    except Exception:
        return items
    try:
        rows = rng.Rows.Count
        cols = rng.Columns.Count
    except Exception:
        return items
    # 限流：最多扫 5000 个 cell
    total = rows * cols
    if total > 5000:
        return [{"warning": f"UsedRange {total} 个单元格超过扫描上限 5000，请传 sample_range 收缩范围"}]
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            try:
                cell = rng.Cells(r, c)
                v = cell.Validation
                try:
                    v_type = int(v.Type)
                except Exception:
                    continue
                if v_type == -4142:  # xlValidateInputOnly 之外的"无"
                    continue
                addr = cell.Address
                # 用 formula1+type 去重，多个相同规则的合并
                key = (v_type, getattr(v, "Formula1", "") or "", getattr(v, "Formula2", "") or "")
                if key in seen:
                    continue
                seen.add(key)
                vtype_map = {1: "whole_number", 2: "decimal", 3: "list", 4: "date",
                             5: "time", 6: "text_length", 7: "custom_formula"}
                info = {
                    "first_cell": addr,
                    "validation_type": vtype_map.get(v_type, f"type_{v_type}"),
                    "formula1": getattr(v, "Formula1", "") or "",
                    "formula2": getattr(v, "Formula2", "") or "",
                }
                items.append(info)
            except Exception:
                pass
    return items


def _apply_cf_format(fc, fmt):
    """模块级条件格式样式应用器（fill_color/font_color/bold/italic）。
    Excel COM 的 FormatCondition.Interior.Color 必须搭配 Pattern=xlSolid(1)
    才能保存到 xlsx，否则颜色信息在保存时会被 Excel 当作未初始化丢弃。
    """
    if not isinstance(fmt, dict):
        return
    if fmt.get("fill_color"):
        try:
            ole_color = color_to_ole(fmt["fill_color"])
            try: fc.Interior.PatternColorIndex = -4105
            except Exception: pass
            try: fc.Interior.Pattern = 1
            except Exception: pass
            fc.Interior.Color = ole_color
            try: fc.Interior.TintAndShade = 0
            except Exception: pass
        except Exception:
            pass
    if fmt.get("font_color"):
        try:
            fc.Font.Color = color_to_ole(fmt["font_color"])
            try: fc.Font.TintAndShade = 0
            except Exception: pass
        except Exception:
            pass
    if "bold" in fmt:
        try: fc.Font.Bold = bool(fmt["bold"])
        except Exception: pass
    if "italic" in fmt:
        try: fc.Font.Italic = bool(fmt["italic"])
        except Exception: pass


def _read_excel_print_settings(sheet):
    """读取打印/页面设置"""
    info = {}
    try:
        ps = sheet.PageSetup
        orient = int(ps.Orientation)
        info["orientation"] = "portrait" if orient == 1 else ("landscape" if orient == 2 else f"code_{orient}")
        try: info["paper_size_code"] = int(ps.PaperSize)
        except Exception: pass
        for src, dst in [
            ("top_margin", "TopMargin"), ("bottom_margin", "BottomMargin"),
            ("left_margin", "LeftMargin"), ("right_margin", "RightMargin"),
            ("header_margin", "HeaderMargin"), ("footer_margin", "FooterMargin"),
        ]:
            try: info[src] = float(getattr(ps, dst))
            except Exception: pass
        try: info["zoom"] = ps.Zoom if isinstance(ps.Zoom, int) else None
        except Exception: pass
        try: info["fit_to_pages_wide"] = ps.FitToPagesWide
        except Exception: pass
        try: info["fit_to_pages_tall"] = ps.FitToPagesTall
        except Exception: pass
        try: info["print_area"] = ps.PrintArea or ""
        except Exception: pass
        try: info["print_titles_rows"] = ps.PrintTitleRows or ""
        except Exception: pass
        try: info["print_titles_columns"] = ps.PrintTitleColumns or ""
        except Exception: pass
        try: info["center_horizontally"] = bool(ps.CenterHorizontally)
        except Exception: pass
        try: info["center_vertically"] = bool(ps.CenterVertically)
        except Exception: pass
        try: info["print_gridlines"] = bool(ps.PrintGridlines)
        except Exception: pass
        try: info["print_headings"] = bool(ps.PrintHeadings)
        except Exception: pass
        for src, dst in [
            ("left_header", "LeftHeader"), ("center_header", "CenterHeader"), ("right_header", "RightHeader"),
            ("left_footer", "LeftFooter"), ("center_footer", "CenterFooter"), ("right_footer", "RightFooter"),
        ]:
            try: info[src] = getattr(ps, dst) or ""
            except Exception: pass
    except Exception as e:
        info["error"] = f"读取页面设置失败: {e}"
    # 标签颜色 / 缩放 / 网格
    try:
        tab = sheet.Tab
        if tab and tab.Color is not None and tab.Color != -4105:
            info["tab_color"] = ole_to_rgb(int(tab.Color))
    except Exception:
        pass
    try:
        info["display_gridlines"] = None  # 需要 ActiveWindow，跳过
    except Exception:
        pass
    return info


def _excel_border_signature(edge_info):
    if not edge_info or edge_info.get("style") == "none":
        return "none"
    color = edge_info.get("color")
    if isinstance(color, (list, tuple)):
        color = ",".join(str(x) for x in color)
    return f"{edge_info.get('style') or ''}|{edge_info.get('weight') or ''}|{color or ''}"


def _build_excel_border_summary(rows):
    edge_counts = {edge: {} for edge in BORDER_EDGE_READ_MAP}
    pattern_counts = {}
    pattern_examples = {}
    bordered_cell_count = 0
    for row in rows:
        for cell_info in row:
            borders = cell_info.get("borders")
            if not borders:
                continue
            edge_sigs = {edge: _excel_border_signature(borders.get(edge)) for edge in BORDER_EDGE_READ_MAP}
            if any(sig != "none" for sig in edge_sigs.values()):
                bordered_cell_count += 1
            for edge, sig in edge_sigs.items():
                edge_counts[edge][sig] = edge_counts[edge].get(sig, 0) + 1
            pattern_key = json.dumps(edge_sigs, sort_keys=True, ensure_ascii=False)
            pattern_counts[pattern_key] = pattern_counts.get(pattern_key, 0) + 1
            if pattern_key not in pattern_examples:
                pattern_examples[pattern_key] = cell_info.get("cell")
    return {
        "has_borders": bordered_cell_count > 0,
        "bordered_cell_count": bordered_cell_count,
        "different_cell_border_patterns": len(pattern_counts),
        "inconsistent_edges": [edge for edge, counts in edge_counts.items() if len(counts) > 1],
        "edge_counts": {
            edge: [
                {"signature": sig, "count": count}
                for sig, count in sorted(counts.items(), key=lambda item: item[1], reverse=True)[:12]
            ]
            for edge, counts in edge_counts.items()
        },
        "cell_pattern_counts": [
            {
                "pattern": json.loads(pattern_key),
                "count": count,
                "example_cell": pattern_examples.get(pattern_key),
            }
            for pattern_key, count in sorted(pattern_counts.items(), key=lambda item: item[1], reverse=True)[:20]
        ],
    }


# ---- 条件格式 / 筛选 读取辅助 ----

# XlFormatConditionType — Excel 条件格式规则类型
_CF_TYPE_NAMES = {
    1: "cell_value",          # xlCellValue
    2: "expression",          # xlExpression
    3: "color_scale",         # xlColorScale
    4: "data_bar",            # xlDatabar
    5: "top10",               # xlTop10
    6: "icon_set",            # xlIconSets
    8: "unique_values",       # xlUniqueValues
    9: "text_string",         # xlTextString
    11: "time_period",        # xlTimePeriod (注意与 xlBlanksCondition 数值冲突，按主流取 xlTimePeriod)
    12: "above_average",      # xlAboveAverageCondition
    16: "errors",             # xlErrorsCondition
    17: "no_errors",          # xlNoErrorsCondition
}

# XlFormatConditionOperator — cell_value 类型规则的运算符
_CF_OPERATOR_NAMES = {
    1: "between",             # xlBetween
    2: "not_between",         # xlNotBetween
    3: "equal",               # xlEqual
    4: "not_equal",           # xlNotEqual
    5: "greater",             # xlGreater
    6: "less",                # xlLess
    7: "greater_equal",       # xlGreaterEqual
    8: "less_equal",          # xlLessEqual
}

# XlAutoFilterOperator — 自动筛选多条件运算符
_FILTER_OPERATOR_NAMES = {
    1: "and",                 # xlAnd
    2: "or",                  # xlOr
    3: "top10_items",         # xlTop10Items
    4: "bottom10_items",      # xlBottom10Items
    5: "top10_percent",       # xlTop10Percent
    6: "bottom10_percent",    # xlBottom10Percent
    7: "filter_values",       # xlFilterValues
    8: "filter_cell_color",   # xlFilterCellColor
    9: "filter_font_color",   # xlFilterFontColor
    10: "filter_icon",        # xlFilterIcon
    11: "filter_dynamic",     # xlFilterDynamic
}


def _read_excel_conditional_formats(rng):
    """读取 Range 上应用的所有条件格式规则。返回 list[dict]。
    每条规则包含：index、type/type_name、applies_to、operator/operator_name、formula1/formula2、format（fill_color/font_color/font_bold/italic/strikethrough）。
    任何子字段读取失败都安全跳过，不会让整体返回空。
    """
    rules = []
    try:
        fc_count = int(rng.FormatConditions.Count)
    except Exception:
        return rules
    for i in range(1, fc_count + 1):
        try:
            cf = rng.FormatConditions(i)
        except Exception:
            continue
        rule = {"index": i}
        # type
        try:
            t = int(cf.Type)
            rule["type"] = t
            rule["type_name"] = _CF_TYPE_NAMES.get(t, f"unknown_{t}")
        except Exception:
            pass
        # applies_to（规则真正覆盖的范围）
        try:
            applies_to = cf.AppliesTo
            rule["applies_to"] = str(applies_to.Address)
        except Exception:
            pass
        # operator（仅 cell_value 类型有意义）
        try:
            op = int(cf.Operator)
            if op > 0:
                rule["operator"] = op
                rule["operator_name"] = _CF_OPERATOR_NAMES.get(op, f"unknown_{op}")
        except Exception:
            pass
        # 公式
        try:
            f1 = cf.Formula1
            if f1 not in (None, ""):
                rule["formula1"] = f1
        except Exception:
            pass
        try:
            f2 = cf.Formula2
            if f2 not in (None, ""):
                rule["formula2"] = f2
        except Exception:
            pass
        # 格式信息（仅对常规规则有效，色阶/数据条/图标集会失败被跳过）
        fmt = {}
        try:
            interior_color = cf.Interior.Color
            if interior_color is not None and interior_color >= 0:
                fmt["fill_color"] = ole_to_rgb(interior_color)
        except Exception:
            pass
        try:
            font_color = cf.Font.Color
            if font_color is not None and font_color >= 0:
                fmt["font_color"] = ole_to_rgb(font_color)
        except Exception:
            pass
        try:
            fmt["font_bold"] = bool(cf.Font.Bold)
        except Exception:
            pass
        try:
            fmt["font_italic"] = bool(cf.Font.Italic)
        except Exception:
            pass
        try:
            fmt["font_strikethrough"] = bool(cf.Font.Strikethrough)
        except Exception:
            pass
        if fmt:
            rule["format"] = fmt
        rules.append(rule)
    return rules


def _read_excel_filters(sheet):
    """读取工作表自动筛选状态。返回 dict 或 None（无筛选）。"""
    try:
        filter_mode = bool(sheet.AutoFilterMode)
    except Exception:
        filter_mode = False
    if not filter_mode:
        return None
    try:
        af = sheet.AutoFilter
        if af is None:
            return None
        info = {"enabled": True}
        try:
            info["range"] = str(af.Range.Address)
        except Exception:
            pass
        try:
            filters_count = int(af.Filters.Count)
        except Exception:
            filters_count = 0
        active_filters = []
        for i in range(1, filters_count + 1):
            try:
                flt = af.Filters(i)
                if not bool(flt.On):
                    continue
                entry = {"column_index": i}
                try:
                    entry["criteria1"] = flt.Criteria1
                except Exception:
                    pass
                try:
                    op = int(flt.Operator)
                    if op > 0:
                        entry["operator"] = op
                        entry["operator_name"] = _FILTER_OPERATOR_NAMES.get(op, f"unknown_{op}")
                except Exception:
                    pass
                try:
                    entry["criteria2"] = flt.Criteria2
                except Exception:
                    pass
                # Criteria1 可能是 tuple（多值筛选）
                if isinstance(entry.get("criteria1"), tuple):
                    entry["criteria1"] = list(entry["criteria1"])
                if isinstance(entry.get("criteria2"), tuple):
                    entry["criteria2"] = list(entry["criteria2"])
                active_filters.append(entry)
            except Exception:
                continue
        info["active_filters"] = active_filters
        info["active_filter_count"] = len(active_filters)
        return info
    except Exception:
        return {"enabled": True, "error": "读取 AutoFilter 详情失败"}


def _base64_char_count(blob):
    return ((len(blob) + 2) // 3) * 4


def _replace_file_extension(name, new_ext):
    root, _ = os.path.splitext(name or "")
    return (root or name or "image") + new_ext


def _flatten_to_rgb(image):
    if image.mode == "RGB":
        return image
    if image.mode in ("RGBA", "LA") or (image.mode == "P" and "transparency" in image.info):
        rgba = image.convert("RGBA")
        background = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        return Image.alpha_composite(background, rgba).convert("RGB")
    return image.convert("RGB")


def _resize_image(image, scale):
    if scale >= 0.999:
        return image.copy()
    width = max(1, int(round(image.width * scale)))
    height = max(1, int(round(image.height * scale)))
    if width == image.width and height == image.height:
        return image.copy()
    return image.resize((width, height), Image.Resampling.LANCZOS)


def _save_png_candidate(image, colors=None):
    candidate = image
    if colors and image.mode != "P":
        try:
            candidate = image.quantize(colors=colors)
        except Exception:
            candidate = image
    out = io.BytesIO()
    candidate.save(out, format="PNG", optimize=True, compress_level=9)
    return out.getvalue()


def _save_jpeg_candidate(image, quality):
    rgb_image = _flatten_to_rgb(image)
    out = io.BytesIO()
    rgb_image.save(out, format="JPEG", quality=quality, optimize=True, progressive=True, subsampling=1)
    return out.getvalue()


def _compress_image_blob_for_inline(name, blob, mime_type):
    original_base64_chars = _base64_char_count(blob)
    if original_base64_chars <= MAX_INLINE_IMAGE_BASE64_CHARS:
        return name, blob, mime_type, {}
    try:
        with Image.open(io.BytesIO(blob)) as opened_image:
            opened_image.load()
            source_image = opened_image.copy()
    except Exception:
        return name, blob, mime_type, {
            "base64_char_count": original_base64_chars,
        }
    best_name = name
    best_blob = blob
    best_mime_type = mime_type
    best_base64_chars = original_base64_chars
    prefer_png = (mime_type or "").lower() == "image/png"
    for scale in IMAGE_SCALE_STEPS:
        resized_image = _resize_image(source_image, scale)
        if prefer_png:
            for colors in PNG_COLOR_STEPS:
                try:
                    png_blob = _save_png_candidate(resized_image, colors)
                except Exception:
                    continue
                png_base64_chars = _base64_char_count(png_blob)
                if png_base64_chars < best_base64_chars:
                    best_name = _replace_file_extension(name, ".png")
                    best_blob = png_blob
                    best_mime_type = "image/png"
                    best_base64_chars = png_base64_chars
                if png_base64_chars <= MAX_INLINE_IMAGE_BASE64_CHARS:
                    return best_name, best_blob, best_mime_type, {
                        "compressed": True,
                        "original_bytes": len(blob),
                        "original_base64_char_count": original_base64_chars,
                        "base64_char_count": best_base64_chars,
                    }
        for quality in JPEG_QUALITY_STEPS:
            try:
                jpeg_blob = _save_jpeg_candidate(resized_image, quality)
            except Exception:
                continue
            jpeg_base64_chars = _base64_char_count(jpeg_blob)
            if jpeg_base64_chars < best_base64_chars:
                best_name = _replace_file_extension(name, ".jpg")
                best_blob = jpeg_blob
                best_mime_type = "image/jpeg"
                best_base64_chars = jpeg_base64_chars
            if jpeg_base64_chars <= MAX_INLINE_IMAGE_BASE64_CHARS:
                return best_name, best_blob, best_mime_type, {
                    "compressed": True,
                    "original_bytes": len(blob),
                    "original_base64_char_count": original_base64_chars,
                    "base64_char_count": best_base64_chars,
                }
    extra = {
        "base64_char_count": best_base64_chars,
    }
    if best_blob != blob or best_mime_type != mime_type or best_name != name:
        extra["compressed"] = True
        extra["original_bytes"] = len(blob)
        extra["original_base64_char_count"] = original_base64_chars
    return best_name, best_blob, best_mime_type, extra


def _build_image_entry(name, blob, mime_type=None, **extra):
    mime_type = mime_type or mimetypes.guess_type(name)[0] or "application/octet-stream"
    name, blob, mime_type, compress_extra = _compress_image_blob_for_inline(name, blob, mime_type)
    item = {
        "name": name,
        "mime_type": mime_type,
        "base64": base64.b64encode(blob).decode("ascii"),
    }
    for key, value in compress_extra.items():
        if value is not None:
            item[key] = value
    for key, value in extra.items():
        if value is not None:
            item[key] = value
    return item


def _extract_images_from_zip(package_path, media_prefix):
    images = []
    if not package_path or not os.path.exists(package_path):
        return images
    try:
        with zipfile.ZipFile(package_path, "r") as zf:
            for name in zf.namelist():
                if not name.startswith(media_prefix):
                    continue
                if name.endswith("/"):
                    continue
                mime_type = mimetypes.guess_type(name)[0] or ""
                if not mime_type.startswith("image/"):
                    continue
                try:
                    blob = zf.read(name)
                except Exception:
                    continue
                images.append(_build_image_entry(os.path.basename(name), blob, mime_type))
    except Exception:
        return []
    return images


def _attach_images(result, images):
    if images:
        result["images"] = images
        result["image_count"] = len(images)
    return result


def _make_temp_copy_path(suffix):
    fd, temp_path = tempfile.mkstemp(suffix=suffix or ".tmp")
    os.close(fd)
    return temp_path


def _safe_remove(path):
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except Exception:
            pass


def _is_http_url(value):
    return bool(re.match(r"^https?://", str(value or "").strip(), re.I))


def _normalize_mime_type(value):
    return str(value or "").split(";")[0].strip().lower()


def _guess_image_extension_from_mime_type(mime_type):
    return REMOTE_IMAGE_CONTENT_TYPE_EXTENSION_MAP.get(_normalize_mime_type(mime_type), ".tmp")


def _ensure_downloads_dir():
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(downloads_dir, exist_ok=True)
    return downloads_dir


def _sanitize_download_name(value):
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", str(value or "").strip()).strip("._") or "image"


def _build_office_download_path(source, suffix):
    parsed_url = urllib.parse.urlparse(source)
    decoded_name = os.path.basename(urllib.parse.unquote(parsed_url.path or "")) or "image"
    base_name = _sanitize_download_name(os.path.splitext(decoded_name)[0])
    ext = suffix or os.path.splitext(decoded_name)[1] or ".tmp"
    nonce = uuid.uuid4().hex[:8]
    return os.path.join(_ensure_downloads_dir(), f"world_office_image_{base_name}_{nonce}{ext}")


def _prepare_office_image_path(image_source):
    source = str(image_source or "").strip()
    if not source:
        return None, None, "缺少图片 path"
    if not _is_http_url(source):
        expanded_path = os.path.expanduser(source)
        if not os.path.exists(expanded_path):
            return None, None, f"图片文件不存在: {source}"
        return expanded_path, None, None

    download_path = None
    try:
        request = urllib.request.Request(
            source,
            headers={"User-Agent": "WorldOfficeTool/1.0"},
        )
        with urllib.request.urlopen(request, timeout=REMOTE_IMAGE_DOWNLOAD_TIMEOUT_SECONDS) as response:
            content_type = _normalize_mime_type(response.headers.get("Content-Type"))
            parsed_url = urllib.parse.urlparse(source)
            suffix = os.path.splitext(urllib.parse.unquote(parsed_url.path or ""))[1]
            if not suffix:
                suffix = _guess_image_extension_from_mime_type(content_type)
            download_path = _build_office_download_path(source, suffix or ".tmp")
            with open(download_path, "wb") as file_obj:
                while True:
                    chunk = response.read(65536)
                    if not chunk:
                        break
                    file_obj.write(chunk)
        if not download_path or not os.path.exists(download_path) or os.path.getsize(download_path) <= 0:
            _safe_remove(download_path)
            return None, None, f"下载公网图片失败: {source}"
        return download_path, None, None
    except Exception as exc:
        _safe_remove(download_path)
        return None, None, f"下载公网图片失败: {source} ({exc})"


def _extract_excel_chart_images_from_workbook(wb):
    """
    提取 Excel 工作簿中所有图表的图片。
    覆盖两类图表：
      1) 嵌在 Worksheet 上的 ChartObjects（sheet.ChartObjects()）
      2) 独立的 Chart Sheet（wb.Charts 集合）
    关键修复：
      - Chart.Export 仅在所在 sheet 处于 Active 时才稳定工作；
        因此对每个含图表的 sheet 先 Activate 再导出，最后还原原 active sheet。
      - 关闭 ScreenUpdating，避免在用户 Excel 上闪屏。
      - 严格校验导出后文件存在且大小 > 0，防止生成空 PNG 当成功。
    """
    images = []
    excel_app = None
    original_active = None
    original_screen_updating = None
    try:
        excel_app = wb.Application
    except Exception:
        excel_app = None
    try:
        original_active = wb.ActiveSheet
    except Exception:
        original_active = None
    if excel_app is not None:
        try:
            original_screen_updating = bool(excel_app.ScreenUpdating)
            excel_app.ScreenUpdating = False
        except Exception:
            original_screen_updating = None

    try:
        # ---- 1) 嵌入式图表（每个 Worksheet 上的 ChartObjects）----
        for sheet in wb.Worksheets:
            try:
                chart_objects = sheet.ChartObjects()
                chart_count = int(chart_objects.Count)
            except Exception:
                continue
            if chart_count <= 0:
                continue
            # 关键：激活当前 sheet，让 Chart.Export 可靠工作
            try:
                sheet.Activate()
            except Exception:
                pass
            for index in range(1, chart_count + 1):
                temp_path = _make_temp_copy_path(".png")
                try:
                    chart_obj = chart_objects(index)
                    try:
                        chart_obj.Chart.Export(temp_path, "PNG")
                    except Exception:
                        # 某些 Excel 版本不接受 FilterName 参数，回落到默认
                        try:
                            chart_obj.Chart.Export(temp_path)
                        except Exception:
                            continue
                    if not os.path.exists(temp_path) or os.path.getsize(temp_path) <= 0:
                        continue
                    with open(temp_path, "rb") as f:
                        blob = f.read()
                    images.append(_build_image_entry(
                        f"{sheet.Name}_chart_{index}.png",
                        blob,
                        "image/png",
                        sheet=sheet.Name,
                        chart_index=index,
                        chart_name=getattr(chart_obj, "Name", None),
                    ))
                except Exception:
                    pass
                finally:
                    _safe_remove(temp_path)

        # ---- 2) 独立 Chart Sheets（wb.Charts 集合，不在 Worksheets 内）----
        try:
            chart_sheet_count = int(wb.Charts.Count)
        except Exception:
            chart_sheet_count = 0
        for index in range(1, chart_sheet_count + 1):
            temp_path = _make_temp_copy_path(".png")
            try:
                chart_sheet = wb.Charts(index)
                try:
                    chart_sheet.Activate()
                except Exception:
                    pass
                try:
                    chart_sheet.Export(temp_path, "PNG")
                except Exception:
                    try:
                        chart_sheet.Export(temp_path)
                    except Exception:
                        continue
                if not os.path.exists(temp_path) or os.path.getsize(temp_path) <= 0:
                    continue
                with open(temp_path, "rb") as f:
                    blob = f.read()
                images.append(_build_image_entry(
                    f"{getattr(chart_sheet, 'Name', f'chart_sheet_{index}')}.png",
                    blob,
                    "image/png",
                    chart_sheet=getattr(chart_sheet, "Name", None),
                    chart_index=index,
                ))
            except Exception:
                pass
            finally:
                _safe_remove(temp_path)
    finally:
        # 还原原 active sheet 和 ScreenUpdating，确保用户视觉无感
        if original_active is not None:
            try:
                original_active.Activate()
            except Exception:
                pass
        if excel_app is not None and original_screen_updating is not None:
            try:
                excel_app.ScreenUpdating = original_screen_updating
            except Exception:
                pass
    return images


def _extract_excel_images_from_workbook(wb):
    suffix = os.path.splitext(getattr(wb, "FullName", "") or getattr(wb, "Name", ""))[1] or ".xlsx"
    temp_path = _make_temp_copy_path(suffix)
    try:
        wb.SaveCopyAs(temp_path)
        images = _extract_images_from_zip(temp_path, "xl/media/")
    except Exception:
        images = []
    finally:
        _safe_remove(temp_path)
    images.extend(_extract_excel_chart_images_from_workbook(wb))
    return images


def _inline_shape_to_png_bytes(inline_shape):
    emf_bits = getattr(inline_shape.Range, "EnhMetaFileBits", None)
    if not emf_bits:
        return None
    image = Image.open(io.BytesIO(bytes(emf_bits)))
    out = io.BytesIO()
    image.save(out, format="PNG")
    return out.getvalue()


def _extract_word_images_from_com(doc):
    images = []
    for index in range(1, doc.InlineShapes.Count + 1):
        try:
            inline_shape = doc.InlineShapes(index)
            blob = _inline_shape_to_png_bytes(inline_shape)
            if not blob:
                continue
            images.append(_build_image_entry(
                f"inline_shape_{index}.png",
                blob,
                "image/png",
                inline_shape_index=index,
                width=round(float(inline_shape.Width), 1) if getattr(inline_shape, "Width", None) is not None else None,
                height=round(float(inline_shape.Height), 1) if getattr(inline_shape, "Height", None) is not None else None,
            ))
        except Exception:
            continue
    return images


def _extract_word_images_from_document(doc):
    full_path = getattr(doc, "FullName", "") or ""
    ext = os.path.splitext(full_path)[1].lower()
    if ext in (".docx", ".docm", ".dotx", ".dotm"):
        images = _extract_images_from_zip(full_path, "word/media/")
        if images:
            return images
    return _extract_word_images_from_com(doc)


# =====================================================
# Excel 操作
# =====================================================

def find_excel_workbook(excel_app, filename):
    """根据文件名查找已打开的工作簿"""
    for wb in excel_app.Workbooks:
        wb_name = wb.Name
        wb_full = wb.FullName
        if filename == wb_name or filename == wb_full:
            return wb
        # 模糊匹配：不带扩展名
        base = os.path.splitext(wb_name)[0]
        if filename == base:
            return wb
        # 路径中包含
        if filename.replace("/", "\\") in wb_full.replace("/", "\\"):
            return wb
    return None


def excel_list_sheets(params):
    """列出工作簿的所有工作表"""
    excel = win32.GetActiveObject("Excel.Application")
    wb = find_excel_workbook(excel, params["filename"])
    if not wb:
        return {"error": f"未找到工作簿: {params['filename']}，当前打开的工作簿: {[w.Name for w in excel.Workbooks]}"}
    sheets = []
    for i in range(1, wb.Sheets.Count + 1):
        s = wb.Sheets(i)
        sheets.append({
            "index": i,
            "name": s.Name,
            "visible": bool(s.Visible),
            "used_range": s.UsedRange.Address if s.UsedRange else "",
        })
    return {"success": True, "workbook": wb.Name, "sheets": sheets}


def excel_read(params):
    """读取 Excel 内容（全量或指定范围）

    分块/限流参数（防止上下文爆炸 & COM 调用超时）：
      - max_rows: 最大返回行数（含 style 默认 200，不含 style 默认 500）
      - max_cells: 最大返回单元格数（max_rows × cols 不超过此值；默认 style=2000 / 无style=8000）
      - head_rows + tail_rows: 首尾分块模式——返回前 head_rows 行 + 后 tail_rows 行，中间省略
      - 截断时返回 truncated:true + total_rows + total_cols + returned_rows + hint
    """
    excel = win32.GetActiveObject("Excel.Application")
    wb = find_excel_workbook(excel, params["filename"])
    if not wb:
        return {"error": f"未找到工作簿: {params['filename']}，当前打开的工作簿: {[w.Name for w in excel.Workbooks]}"}
    images = _extract_excel_images_from_workbook(wb)

    sheet_name = params.get("sheet")
    if sheet_name:
        try:
            sheet = wb.Sheets(sheet_name)
        except:
            return {"error": f"工作表 '{sheet_name}' 不存在，可用: {[wb.Sheets(i).Name for i in range(1, wb.Sheets.Count+1)]}"}
    else:
        sheet = wb.ActiveSheet

    range_addr = params.get("range")
    include_style = params.get("include_style", False)

    if range_addr:
        rng = sheet.Range(range_addr)
    else:
        rng = sheet.UsedRange

    if rng is None:
        return _attach_images({"success": True, "sheet": sheet.Name, "data": [], "rows": 0, "cols": 0}, images)

    # ---------- 分块保护：限流到 max_rows / max_cells ----------
    # 评估原始范围的行列数
    try:
        total_rows = int(rng.Rows.Count)
        total_cols = int(rng.Columns.Count)
    except Exception:
        total_rows, total_cols = 0, 0
    original_address = rng.Address

    DEFAULT_MAX_CELLS = 2000 if include_style else 8000
    DEFAULT_MAX_ROWS = 200 if include_style else 500

    def _to_int_or_none(v):
        try:
            n = int(v)
            return n if n > 0 else None
        except (TypeError, ValueError):
            return None

    head_rows_arg = _to_int_or_none(params.get("head_rows"))
    tail_rows_arg = _to_int_or_none(params.get("tail_rows"))
    max_rows_arg = _to_int_or_none(params.get("max_rows"))
    max_cells_arg = _to_int_or_none(params.get("max_cells"))

    truncated_info = None
    skipped_rows_meta = None  # head+tail 模式下中间省略的描述

    cols_for_calc = max(total_cols, 1)

    # 计算"安全行数上限"
    if max_rows_arg is not None:
        safe_rows = max_rows_arg
    elif max_cells_arg is not None:
        safe_rows = max(1, max_cells_arg // cols_for_calc)
    else:
        safe_rows = min(DEFAULT_MAX_ROWS, max(1, DEFAULT_MAX_CELLS // cols_for_calc))

    def col_letter(col_num):
        """列号(1-based) → 列字母(A, B, ..., Z, AA, AB, ...)"""
        s = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            s = chr(65 + remainder) + s
        return s

    # ---------- 内部函数：从一个连续 Range 提取行数据 ----------
    seen_merge_areas = set()
    merge_ranges = []

    def _collect_block(target_rng):
        """读取一个连续 Range，返回 (rows, column_letters)。merge_ranges 由外层 set 去重。"""
        block_values = target_rng.Value
        if block_values is None:
            return [], []
        if not isinstance(block_values, tuple):
            block_values = ((block_values,),)
        if block_values and not isinstance(block_values[0], tuple):
            block_values = (block_values,)
        b_start_row = target_rng.Row
        b_start_col = target_rng.Column
        b_col_count = len(block_values[0]) if block_values else 0
        b_column_letters = [col_letter(b_start_col + i) for i in range(b_col_count)]
        b_rows = []
        for r_idx, row in enumerate(block_values):
            row_data = []
            for c_idx, val in enumerate(row):
                if isinstance(val, int) and val in EXCEL_ERROR_MAP:
                    val = EXCEL_ERROR_MAP[val]
                if isinstance(val, Decimal):
                    val = float(val)
                addr = f"{b_column_letters[c_idx]}{b_start_row + r_idx}"
                cell_info = {"cell": addr, "value": val}
                cell = sheet.Cells(b_start_row + r_idx, b_start_col + c_idx)
                merged = bool(cell.MergeCells)
                cell_info["merged"] = merged
                if merged:
                    merge_area = cell.MergeArea
                    merge_anchor = merge_area.Cells(1, 1)
                    merge_area_addr = merge_area.Address
                    merge_anchor_addr = merge_anchor.Address
                    is_merge_anchor = bool(
                        cell.Row == merge_anchor.Row and cell.Column == merge_anchor.Column
                    )
                    cell_info["merge_area"] = merge_area_addr
                    cell_info["merge_anchor"] = merge_anchor_addr
                    cell_info["is_merge_anchor"] = is_merge_anchor
                    cell_info["writable"] = is_merge_anchor
                    cell_info["write_target"] = merge_anchor_addr
                    if not is_merge_anchor:
                        cell_info["merged_shadow"] = True
                    if merge_area_addr not in seen_merge_areas:
                        seen_merge_areas.add(merge_area_addr)
                        merge_ranges.append({
                            "range": merge_area_addr,
                            "anchor": merge_anchor_addr,
                        })
                else:
                    cell_info["writable"] = True
                    cell_info["write_target"] = addr
                if isinstance(val, str) and val.strip():
                    try:
                        float(val)
                        if cell.NumberFormat == '@':
                            cell_info["text_number"] = True
                    except (ValueError, TypeError):
                        pass
                if include_style:
                    cell_info["formula"] = cell.Formula if cell.HasFormula else None
                    font = cell.Font
                    cell_info["font"] = {
                        "name": font.Name,
                        "size": font.Size,
                        "bold": bool(font.Bold),
                        "italic": bool(font.Italic),
                        "color": ole_to_rgb(font.Color) if font.Color else None,
                        "underline": font.Underline,
                    }
                    cell_info["fill_color"] = ole_to_rgb(cell.Interior.Color) if cell.Interior.Pattern != -4142 else None
                    cell_info["h_align"] = cell.HorizontalAlignment
                    cell_info["v_align"] = cell.VerticalAlignment
                    cell_info["wrap_text"] = bool(cell.WrapText)
                    cell_info["number_format"] = cell.NumberFormat
                    cell_info["borders"] = _read_excel_cell_borders(cell)
                row_data.append(cell_info)
            b_rows.append(row_data)
        return b_rows, b_column_letters

    effective_range_address = rng.Address  # 默认为原 Range 地址；下面被截断/拼接时会覆盖

    # —— head + tail 分块模式 ——（分两段独立读取后拼接，避免 MultiArea .Value 只取第一段）
    if (head_rows_arg or tail_rows_arg) and total_rows > 0:
        head_n = head_rows_arg or 0
        tail_n = tail_rows_arg or 0
        if head_n + tail_n >= total_rows:
            # 首尾相加 ≥ 总行数，直接全读，不截断
            rows, column_letters = _collect_block(rng)
        else:
            head_rng = rng.Resize(head_n, total_cols) if head_n > 0 else None
            tail_rng = (
                sheet.Range(
                    sheet.Cells(rng.Row + total_rows - tail_n, rng.Column),
                    sheet.Cells(rng.Row + total_rows - 1, rng.Column + total_cols - 1),
                )
                if tail_n > 0
                else None
            )
            head_rows_data, head_cols = ([], [])
            tail_rows_data, tail_cols = ([], [])
            if head_rng is not None:
                head_rows_data, head_cols = _collect_block(head_rng)
            if tail_rng is not None:
                tail_rows_data, tail_cols = _collect_block(tail_rng)
            column_letters = head_cols or tail_cols
            effective_range_address = "+".join(filter(None, [
                head_rng.Address if head_rng is not None else None,
                tail_rng.Address if tail_rng is not None else None,
            ]))
            skipped_n = max(0, total_rows - head_n - tail_n)
            # 中间插一个 separator row（占 1 个对象，体积可忽略）
            separator_row = [{
                "separator": True,
                "skipped_rows": skipped_n,
                "skipped_range": (
                    f"{column_letters[0]}{rng.Row + head_n}:"
                    f"{column_letters[-1] if column_letters else ''}"
                    f"{rng.Row + total_rows - tail_n - 1}"
                    if column_letters and skipped_n > 0
                    else None
                ),
                "hint": f"中间省略 {skipped_n} 行，如需查看请用 range 参数读取该子范围",
            }] if skipped_n > 0 else []
            rows = head_rows_data + ([separator_row] if separator_row else []) + tail_rows_data
            truncated_info = {
                "truncated": True,
                "mode": "head_tail",
                "original_address": original_address,
                "total_rows": total_rows,
                "total_cols": total_cols,
                "returned_rows": head_n + tail_n,
                "skipped_rows": skipped_n,
                "head_rows": head_n,
                "tail_rows": tail_n,
                "hint": (
                    f"已采用 head+tail 模式：返回前 {head_n} 行 + 后 {tail_n} 行，"
                    f"中间省略 {skipped_n} 行。如需读取省略部分，请用 range 参数指定具体子范围。"
                ),
            }
    else:
        # —— 单段截断模式 ——（在没有 head/tail 时生效）
        if total_rows > safe_rows:
            base_row = rng.Row
            rng = rng.Resize(safe_rows, total_cols)
            effective_range_address = rng.Address
            next_start = base_row + safe_rows
            last_col_letter = col_letter(rng.Column + total_cols - 1)
            truncated_info = {
                "truncated": True,
                "mode": "head",
                "original_address": original_address,
                "total_rows": total_rows,
                "total_cols": total_cols,
                "returned_rows": safe_rows,
                "hint": (
                    f"原始数据 {total_rows} 行 × {total_cols} 列已截断到前 {safe_rows} 行 "
                    f"（{'含样式' if include_style else '不含样式'}时默认上限 {DEFAULT_MAX_ROWS} 行 / "
                    f"{DEFAULT_MAX_CELLS} 单元格）。如需读取剩余部分，请用 "
                    f"range='{col_letter(rng.Column)}{next_start}:{last_col_letter}{base_row + total_rows - 1}' 继续读取，"
                    f"或一次性用 head_rows + tail_rows 拿到首尾预览。"
                ),
            }
        rows, column_letters = _collect_block(rng)

    result = {
        "success": True,
        "sheet": sheet.Name,
        "range": effective_range_address,
        "columns": column_letters,
        "rows": len(rows),
        "cols": len(rows[0]) if rows else 0,
        "merge_ranges": merge_ranges,
        "data": rows,
    }
    if truncated_info:
        result.update(truncated_info)
    if include_style:
        result["border_summary"] = _build_excel_border_summary(rows)

    # 可选：读取条件格式规则（默认关闭，需显式开启）
    if params.get("include_conditional_formats"):
        try:
            result["conditional_formats"] = _read_excel_conditional_formats(rng)
        except Exception as exc:
            result["conditional_formats"] = []
            result["conditional_formats_error"] = f"读取条件格式失败: {exc}"

    # 可选：读取自动筛选状态（默认关闭，需显式开启）
    if params.get("include_filters"):
        try:
            result["auto_filter"] = _read_excel_filters(sheet)
        except Exception as exc:
            result["auto_filter"] = None
            result["auto_filter_error"] = f"读取自动筛选失败: {exc}"

    # 可选：读取图表清单（关键，决定 AI 能否"看见"已有图表）
    if params.get("include_charts"):
        try:
            include_series = params.get("include_chart_series", True)
            charts = []
            chart_idx = 0
            for i in range(1, sheet.Shapes.Count + 1):
                sh = sheet.Shapes(i)
                if sh.HasChart:
                    chart_idx += 1
                    brief = _read_chart_brief(sh, include_series=include_series)
                    brief["chart_index"] = chart_idx
                    charts.append(brief)
            result["charts"] = charts
            result["chart_count"] = len(charts)
        except Exception as exc:
            result["charts"] = []
            result["charts_error"] = f"读取图表失败: {exc}"

    # 可选：读取 ListObject（Excel 表格 Ctrl+T）
    if params.get("include_list_objects") or params.get("include_tables"):
        try:
            result["list_objects"] = _read_excel_list_objects(sheet)
        except Exception as exc:
            result["list_objects"] = []
            result["list_objects_error"] = f"读取表格失败: {exc}"

    # 可选：读取数据验证规则（开销大，按需打开）
    if params.get("include_data_validations"):
        try:
            sample_range = params.get("validation_sample_range")
            result["data_validations"] = _read_excel_data_validations(sheet, sample_range)
        except Exception as exc:
            result["data_validations"] = []
            result["data_validations_error"] = f"读取数据验证失败: {exc}"

    # 可选：读取形状/图片/文本框（不含图表）
    if params.get("include_shapes"):
        try:
            result["shapes"] = _read_excel_shapes(sheet, max_shapes=int(params.get("shapes_limit", 200)))
            result["shape_count"] = len(result["shapes"])
        except Exception as exc:
            result["shapes"] = []
            result["shapes_error"] = f"读取形状失败: {exc}"

    # 可选：读取超链接
    if params.get("include_hyperlinks"):
        try:
            result["hyperlinks"] = _read_excel_hyperlinks(sheet, max_links=int(params.get("hyperlinks_limit", 500)))
        except Exception as exc:
            result["hyperlinks"] = []
            result["hyperlinks_error"] = f"读取超链接失败: {exc}"

    # 可选：读取打印 / 页面设置
    if params.get("include_print_settings"):
        try:
            result["print_settings"] = _read_excel_print_settings(sheet)
        except Exception as exc:
            result["print_settings"] = None
            result["print_settings_error"] = f"读取打印设置失败: {exc}"

    return _attach_images(result, images)


def excel_edit(params):
    """编辑 Excel（值/公式/样式/条件格式/行列操作等）"""
    excel = win32.GetActiveObject("Excel.Application")
    wb = find_excel_workbook(excel, params["filename"])
    if not wb:
        return {"error": f"未找到工作簿: {params['filename']}"}

    sheet_name = params.get("sheet") or params.get("sheet_name")
    if sheet_name:
        try:
            sheet = wb.Sheets(sheet_name)
        except:
            return {"error": f"工作表 '{sheet_name}' 不存在"}
    else:
        sheet = wb.ActiveSheet

    action = params.get("edit_action", "set_value")
    results = []
    errors = []

    def get_excel_operations(params, required_keys):
        """兼容单操作快捷参数：operations 或直接 cell/value/formula/style 形式"""
        operations = params.get("operations")
        if isinstance(operations, dict):
            operations = [operations]
        if isinstance(operations, list) and len(operations) > 0:
            return operations

        if all(key in params for key in required_keys):
            return [{key: params[key] for key in required_keys}]

        return None

    def get_received_keys_text(params):
        keys = sorted([k for k in params.keys() if k != "tool"])
        return ", ".join(keys) if keys else "无"

    def get_writable_excel_cell(cell_ref):
        cell = sheet.Range(cell_ref)
        if bool(cell.MergeCells):
            merge_area = cell.MergeArea
            merge_anchor = merge_area.Cells(1, 1)
            if cell.Row != merge_anchor.Row or cell.Column != merge_anchor.Column:
                merge_addr = merge_area.Address
                anchor_addr = merge_anchor.Address
                return None, (
                    f"{cell_ref} 位于合并单元格 {merge_addr} 中，但不是左上角可写单元格；"
                    f"请改为写入 {anchor_addr}"
                )
        return cell, None

    # ===== 设置值 =====
    if action == "set_value":
        operations = get_excel_operations(params, ["cell", "value"])
        if not operations:
            operations = get_excel_operations(params, ["range", "value"])
        if not operations:
            return {"error": f"set_value 需要 operations，或直接提供 cell/range + value 参数；收到字段: {get_received_keys_text(params)}"}
        for op in operations:
            cell_ref = op.get("cell") or op.get("range")
            cell, write_error = get_writable_excel_cell(cell_ref)
            if write_error:
                errors.append(write_error)
                results.append(f"⚠️ {write_error}")
                continue
            val = op["value"]
            # 自动将数字字符串转为真正的数字（避免Excel存储为文本）
            if isinstance(val, str):
                try:
                    val = int(val)
                except ValueError:
                    try:
                        val = float(val)
                    except ValueError:
                        pass
            # 如果单元格格式是文本(@)且要写入数字，先改为常规格式
            if isinstance(val, (int, float)) and cell.NumberFormat == '@':
                cell.NumberFormat = "General"
            cell.Value = val
            results.append(f"{cell_ref} = {val}")

    # ===== 设置公式 =====
    elif action == "set_formula":
        operations = get_excel_operations(params, ["cell", "formula"])
        if not operations:
            operations = get_excel_operations(params, ["range", "formula"])
        if not operations:
            return {"error": f"set_formula 需要 operations，或直接提供 cell/range + formula 参数；收到字段: {get_received_keys_text(params)}"}
        for op in operations:
            cell_ref = op.get("cell") or op.get("range")
            cell, write_error = get_writable_excel_cell(cell_ref)
            if write_error:
                errors.append(write_error)
                results.append(f"⚠️ {write_error}")
                continue
            cell.Formula = op["formula"]
            excel.Calculate()
            if not bool(cell.HasFormula):
                formula_error = (
                    f"{cell_ref} 公式未成功写入；如果目标位于合并区域，请读取 merge_anchor 并改写左上角单元格"
                )
                errors.append(formula_error)
                results.append(f"⚠️ {formula_error}")
                continue
            # 验证公式结果，报告错误
            result_val = cell.Value
            if isinstance(result_val, int) and result_val in EXCEL_ERROR_MAP:
                err_name = EXCEL_ERROR_MAP[result_val]
                results.append(f"{cell_ref} 公式 = {op['formula']} → ⚠️{err_name}（公式引用的单元格可能是文本格式，请先用 convert_text_numbers 转换）")
            else:
                results.append(f"{cell_ref} 公式 = {op['formula']} → {result_val}")

    # ===== 文本型数字转换为真正数字 =====
    elif action == "convert_text_numbers":
        target_range = params.get("range")
        if not target_range:
            return {"error": f"convert_text_numbers 需要 range 参数，如 'C4:C13'"}
        rng = sheet.Range(target_range)
        converted = 0
        for cell in rng:
            # 使用 Value2 获取原始值（不经过日期/货币转换）
            val = cell.Value2
            nf = cell.NumberFormat
            is_text_format = (nf == '@')
            if isinstance(val, str) and val.strip():
                try:
                    num = int(val)
                except ValueError:
                    try:
                        num = float(val)
                    except ValueError:
                        continue
                # 先将格式改为常规，再写入数字
                if is_text_format:
                    cell.NumberFormat = "General"
                cell.Value = num
                converted += 1
            elif is_text_format and isinstance(val, (int, float)):
                # 格式是文本但值已经是数字（COM自动转换了），修正格式
                cell.NumberFormat = "General"
                cell.Value = val
                converted += 1
        results.append(f"{target_range} 已转换 {converted} 个文本型数字为真正数字")

    # ===== 设置样式 =====
    elif action == "set_style":
        operations = get_excel_operations(params, ["range", "style"])
        if not operations:
            return {"error": f"set_style 需要 operations，或直接提供 range + style 参数；收到字段: {get_received_keys_text(params)}"}
        for op in operations:
            rng = sheet.Range(op["range"])
            style = op.get("style", {})

            # 字体
            font = style.get("font", {})
            if "name" in font:
                rng.Font.Name = font["name"]
            if "size" in font:
                rng.Font.Size = font["size"]
            if "bold" in font:
                rng.Font.Bold = font["bold"]
            if "italic" in font:
                rng.Font.Italic = font["italic"]
            if font.get("color") is not None:
                rng.Font.Color = color_to_ole(font["color"])
            if "underline" in font:
                rng.Font.Underline = _resolve_excel_underline(font["underline"])
            if "strikethrough" in font:
                rng.Font.Strikethrough = _to_bool(font["strikethrough"], default=False)

            # 对齐
            if "h_align" in style:
                rng.HorizontalAlignment = _resolve_excel_halign(style["h_align"])
            if "v_align" in style:
                rng.VerticalAlignment = _resolve_excel_valign(style["v_align"])

            # 自动换行
            if "wrap_text" in style:
                rng.WrapText = style["wrap_text"]

            # 文字方向/角度
            if "orientation" in style:
                rng.Orientation = style["orientation"]

            # 缩进
            if "indent" in style:
                rng.IndentLevel = style["indent"]

            # 背景色
            if style.get("fill_color") is not None:
                rng.Interior.Color = color_to_ole(style["fill_color"])

            # 数字格式
            if "number_format" in style:
                rng.NumberFormat = style["number_format"]

            # 边框
            borders = style.get("borders", {})
            if borders:
                if isinstance(borders, dict) and "all" in borders:
                    b = borders["all"]
                    for edge in BORDER_ALL_EDGE_INDEXES:
                        try:
                            _apply_excel_border_style(rng.Borders(edge), b)
                        except Exception:
                            pass
                else:
                    for edge_name, b in borders.items():
                        if edge_name == "all":
                            continue
                        edge_idx = BORDER_INDEX_MAP.get(edge_name)
                        if edge_idx is None:
                            continue
                        _apply_excel_border_style(rng.Borders(edge_idx), b)

            # 列宽/行高
            if "column_width" in style:
                rng.ColumnWidth = style["column_width"]
            if "row_height" in style:
                rng.RowHeight = style["row_height"]
            if "auto_fit_column" in style and style["auto_fit_column"]:
                rng.EntireColumn.AutoFit()
            if "auto_fit_row" in style and style["auto_fit_row"]:
                rng.EntireRow.AutoFit()

            results.append(f"{op['range']} 样式已更新")

    # ===== 合并/取消合并 =====
    elif action == "merge":
        rng = sheet.Range(params["range"])
        rng.Merge()
        results.append(f"{params['range']} 已合并")

    elif action == "unmerge":
        rng = sheet.Range(params["range"])
        rng.UnMerge()
        results.append(f"{params['range']} 已取消合并")

    # ===== 插入/删除行 =====
    elif action == "insert_rows":
        row_num = params["row"]
        count = params.get("count", 1)
        for _ in range(count):
            sheet.Rows(row_num).Insert()
        results.append(f"在第{row_num}行前插入了{count}行")

    elif action == "delete_rows":
        row_num = params["row"]
        count = params.get("count", 1)
        sheet.Range(f"{row_num}:{row_num + count - 1}").EntireRow.Delete()
        results.append(f"删除了第{row_num}到{row_num+count-1}行")

    # ===== 插入/删除列 =====
    elif action == "insert_cols":
        col = params["col"]  # 可以是 "A" 或数字
        count = params.get("count", 1)
        for _ in range(count):
            sheet.Columns(col).Insert()
        results.append(f"在{col}列前插入了{count}列")

    elif action == "delete_cols":
        col = params["col"]
        count = params.get("count", 1)
        if isinstance(col, str) and len(col) <= 3:
            # 字母列名
            for _ in range(count):
                sheet.Columns(col).Delete()
        else:
            sheet.Columns(f"{col}:{col}").EntireColumn.Delete()
        results.append(f"删除了{col}列起{count}列")

    # ===== 排序 =====
    elif action == "sort":
        rng = sheet.Range(params["range"])
        has_header = params.get("has_header", True)
        xl_header = 1 if has_header else 2  # xlYes=1, xlNo=2

        # 入参兼容：
        #   旧单列：sort_column:2, order:"desc"
        #   新多列：sort_keys:[{column:2,order:"desc"},{column:9,order:"asc"}]（最多3列，Excel原生支持）
        sort_keys = params.get("sort_keys")
        if not sort_keys:
            sort_keys = [{"column": params.get("sort_column", 1), "order": params.get("order", "asc")}]
        # 容错：sort_keys 是 JSON 字符串
        if isinstance(sort_keys, str):
            try:
                sort_keys = json.loads(sort_keys)
            except Exception:
                sort_keys = [{"column": 1, "order": "asc"}]

        norm_keys = []
        for k in sort_keys:
            if isinstance(k, dict):
                col = k.get("column", k.get("sort_column", 1))
                ordr = k.get("order", "asc")
            elif isinstance(k, (list, tuple)) and len(k) >= 1:
                col = k[0]
                ordr = k[1] if len(k) >= 2 else "asc"
            else:
                col = k
                ordr = "asc"
            try:
                col = int(col)
            except (TypeError, ValueError):
                col = 1
            xl_ord = 2 if str(ordr).strip().lower() in ("desc", "descending", "降序", "z-a", "2") else 1
            norm_keys.append((col, xl_ord))

        if not norm_keys:
            norm_keys = [(1, 1)]

        # Excel 原生 Sort 单次调用最多支持 3 个 Key
        keys_to_use = norm_keys[:3]
        sort_kwargs = {"Header": xl_header}
        for i, (col, xl_ord) in enumerate(keys_to_use, start=1):
            sort_kwargs[f"Key{i}"] = rng.Columns(col)
            sort_kwargs[f"Order{i}"] = xl_ord
        rng.Sort(**sort_kwargs)

        desc = "、".join(f"第{c}列{'升序' if o == 1 else '降序'}" for c, o in keys_to_use)
        results.append(f"{params['range']} 已按 {desc} 排序")

    # ===== 自动筛选 =====
    elif action == "auto_filter":
        rng = sheet.Range(params["range"])
        field = params.get("field")  # 列号（1开始）
        criteria = params.get("criteria")
        criteria2 = params.get("criteria2")
        operator = params.get("operator")  # 可选 xl 操作符（1=And, 2=Or, 7=FilterValues 等）

        def _normalize_filter_criteria(val):
            """把 criteria 规范化为 Excel AutoFilter 能识别的格式
            返回 (criteria1, auto_operator)：
              - list/tuple        → ([str,...], 7)  xlFilterValues 多值
              - "A,B,C" 多值      → ([A,B,C], 7)    自动拆分（仅当不含比较运算符且分段≥2）
              - "[\"A\",\"B\"]" JSON 数组 → 同上
              - 数字              → (str(num), None)
              - "=text" / ">100" / "<>x" / "*文" / "?x" 已带运算符/通配符 → 原样
              - 其他纯文本         → 原样直接传给 Excel（Excel AutoFilter 会做精确匹配）
            ⚠️ 不要把 "A1" 包成 '="A1"'！Excel AutoFilter 会把整个 ="A1" 当字面值，0 行命中。
               条件格式的 formula1 才需要那种包装，AutoFilter 的 criteria 直接传字面值就够。
            """
            if isinstance(val, (list, tuple)):
                items = [str(x).strip() for x in val if str(x).strip() != ""]
                return (items, 7)
            if isinstance(val, bool):
                return (str(val), None)
            if isinstance(val, (int, float)):
                return (str(val), None)
            s = "" if val is None else str(val).strip()
            if not s:
                return (s, None)
            # JSON 数组字符串 ["a","b"] → 多值
            if (s.startswith("[") and s.endswith("]")):
                try:
                    parsed = json.loads(s)
                    if isinstance(parsed, list) and len(parsed) >= 1:
                        return ([str(x).strip() for x in parsed if str(x).strip()], 7)
                except Exception:
                    pass
            # 已带比较运算符 / 通配符 → 原样保留
            if s.startswith(("=", ">", "<")) or "*" in s or "?" in s:
                return (s, None)
            # 含逗号且至少 2 段 → 多值筛选
            if "," in s or "，" in s:
                parts = [p.strip() for p in re.split(r"[,，]", s) if p.strip()]
                if len(parts) >= 2:
                    return (parts, 7)
            # 普通文本：直接传给 Excel，由 Excel AutoFilter 做精确匹配（不要再包 ="..."）
            return (s, None)

        if field is not None and criteria is not None:
            crit1, auto_op = _normalize_filter_criteria(criteria)
            try:
                xl_op = int(operator) if operator is not None else auto_op
            except (TypeError, ValueError):
                xl_op = auto_op
            try:
                field_int = int(field)
            except (TypeError, ValueError):
                field_int = field

            kwargs = {"Field": field_int, "Criteria1": crit1}
            if xl_op is not None:
                kwargs["Operator"] = xl_op
            if criteria2 is not None and xl_op in (1, 2):  # xlAnd / xlOr 才用 Criteria2
                kwargs["Criteria2"] = str(criteria2)
            rng.AutoFilter(**kwargs)
            crit_desc = crit1 if not isinstance(crit1, list) else "/".join(crit1)
            results.append(f"{params['range']} 已筛选：第{field_int}列 = {crit_desc}")
        elif field is None and criteria is None:
            rng.AutoFilter()
            results.append(f"{params['range']} 自动筛选已切换")
        else:
            try:
                field_int = int(field)
            except (TypeError, ValueError):
                field_int = field
            rng.AutoFilter(Field=field_int)
            results.append(f"{params['range']} 第{field_int}列筛选已切换")

    # ===== 清除筛选 =====
    elif action == "clear_filter":
        if sheet.AutoFilterMode:
            sheet.AutoFilterMode = False
        results.append("筛选已清除")

    # ===== 条件格式 =====
    elif action == "conditional_format":
        rng = sheet.Range(params["range"])
        cf_type = params.get("cf_type", "cell_value")

        # 兼容 format 被序列化成 JSON 字符串的情况（模型常误传 '{"fill_color":"浅红"}' 这种字符串）
        fmt = params.get("format", {})
        if isinstance(fmt, str):
            try:
                fmt = json.loads(fmt) if fmt.strip() else {}
            except Exception:
                fmt = {}
        if not isinstance(fmt, dict):
            fmt = {}

        def _apply_cf_format(fc, fmt):
            """统一应用条件格式的样式（fill_color/font_color/bold/italic）。
            关键修复：Excel COM 的 FormatCondition.Interior.Color 必须搭配
            Pattern=xlSolid(1) 才能保存到 xlsx，否则颜色信息在保存时会被
            Excel 当作未初始化丢弃，再读时变成 0（黑色）。
            """
            if not isinstance(fmt, dict):
                return
            if fmt.get("fill_color"):
                try:
                    ole_color = color_to_ole(fmt["fill_color"])
                    # 必须先设 Pattern=xlSolid，颜色才会被持久化
                    try:
                        fc.Interior.PatternColorIndex = -4105  # xlAutomatic（部分版本要求）
                    except Exception:
                        pass
                    try:
                        fc.Interior.Pattern = 1  # xlSolid
                    except Exception:
                        pass
                    fc.Interior.Color = ole_color
                    try:
                        fc.Interior.TintAndShade = 0  # 防主题色干扰
                    except Exception:
                        pass
                except Exception:
                    pass
            if fmt.get("font_color"):
                try:
                    fc.Font.Color = color_to_ole(fmt["font_color"])
                    try:
                        fc.Font.TintAndShade = 0
                    except Exception:
                        pass
                except Exception:
                    pass
            if "bold" in fmt:
                try:
                    fc.Font.Bold = bool(fmt["bold"])
                except Exception:
                    pass
            if "italic" in fmt:
                try:
                    fc.Font.Italic = bool(fmt["italic"])
                except Exception:
                    pass

        def _wrap_text_formula(val):
            """文本相等比较时，formula1 必须是 Excel 表达式 ="文本"
            自动包装规则：
              - None/空 → ""
              - 已 = 开头 → 原样返回（允许模型自己写好表达式）
              - 纯数字 → 原样返回
              - 文本 → 去掉外层引号后，包成 ="..."（内部双引号 → ""）
            """
            if val is None:
                return ""
            s = str(val).strip()
            if not s:
                return ""
            if s.startswith("="):
                return s
            try:
                float(s)
                return s
            except ValueError:
                pass
            if len(s) >= 2 and s.startswith('"') and s.endswith('"'):
                s = s[1:-1]
            return '="' + s.replace('"', '""') + '"'

        if cf_type == "cell_value":
            # xlCellValue = 1
            # operator: xlBetween=1, xlNotBetween=2, xlEqual=3, xlNotEqual=4,
            #          xlGreater=5, xlLess=6, xlGreaterEqual=7, xlLessEqual=8
            try:
                operator = int(params.get("operator", 5))
            except (TypeError, ValueError):
                operator = 5
            formula1 = params.get("formula1", "")
            formula2 = params.get("formula2")

            # === 关键修复：文本相等/不等 → 自动转 xlExpression 公式模式 ===
            # Excel COM 的 FormatConditions.Add(xlCellValue, xlEqual, ="中文") 在某些
            # Excel 版本下会直接抛 E_INVALIDARG 异常（xlCellValue+xlEqual+中文文本是
            # known 不稳定路径）。改用 xlExpression(2) 公式模式 =$N2="业绩差" 100% 可靠。
            use_formula_fallback = False
            text_value = None
            if operator in (3, 4):
                t = "" if formula1 is None else str(formula1).strip()
                if t.startswith("="):
                    t = t[1:].strip()
                if len(t) >= 2 and t.startswith('"') and t.endswith('"'):
                    t = t[1:-1]
                # 是数字吗？数字走原 cell_value 路径；文本走 formula fallback
                try:
                    float(t)
                except (TypeError, ValueError):
                    if t:
                        use_formula_fallback = True
                        text_value = t

            if use_formula_fallback:
                # xlExpression: 用 =$AnchorCol{Row}="文本" 表达式
                # 注意：win32com 早期绑定(makepy)下 Range.Address 是只读字符串属性，
                # 不能 Address(False, False) 调用，否则 "'str' object is not callable"。
                # 用无参 .Address 拿绝对地址 "$N$2" 再正则解析，兼容早期/晚期绑定。
                anchor_addr_abs = rng.Cells(1, 1).Address  # "$N$2" 早期/晚期都返回字符串
                m_addr = re.match(r"\$?([A-Z]+)\$?(\d+)", str(anchor_addr_abs))
                if m_addr:
                    anchor_ref = f"${m_addr.group(1)}{m_addr.group(2)}"  # $N2 列绝对、行相对
                else:
                    anchor_ref = str(anchor_addr_abs)
                op_str = "=" if operator == 3 else "<>"
                text_escaped = text_value.replace('"', '""')
                expr = f'={anchor_ref}{op_str}"{text_escaped}"'
                # 关键修复：win32com 晚期绑定下用关键字参数 Formula1=expr 跳过 Optional
                # 参数 Operator 不可靠，Excel 端会收到错乱的参数抛 DISP_E_PARAMNOTFOUND。
                # 必须用纯位置参数 + None 占位 Operator（None → VT_NULL，Excel 视为缺省）。
                fc = rng.FormatConditions.Add(2, None, expr)  # xlExpression
                _apply_cf_format(fc, fmt)
                results.append(f"{params['range']} 条件格式(单元格值{op_str}{text_value!r})已添加[via formula]")
            else:
                # 原 cell_value 路径：数字比较、between/greater/less 等
                if operator in (3, 4):
                    formula1 = _wrap_text_formula(formula1)
                    if formula2:
                        formula2 = _wrap_text_formula(formula2)
                if formula2:
                    fc = rng.FormatConditions.Add(1, operator, formula1, formula2)
                else:
                    fc = rng.FormatConditions.Add(1, operator, formula1)
                _apply_cf_format(fc, fmt)
                results.append(f"{params['range']} 条件格式(单元格值)已添加")

        elif cf_type == "formula":
            # xlExpression = 2
            formula = params.get("formula") or params.get("formula1") or ""
            if formula and not str(formula).startswith("="):
                formula = "=" + str(formula)
            # 关键修复：用位置参数 + None 占位 Operator，避免晚期绑定下关键字参数失效
            fc = rng.FormatConditions.Add(2, None, formula)
            _apply_cf_format(fc, fmt)
            results.append(f"{params['range']} 条件格式(公式)已添加")

        elif cf_type == "databar":
            fc = rng.FormatConditions.AddDatabar()
            if "color" in params:
                fc.BarColor.Color = color_to_ole(params["color"])
            elif fmt.get("fill_color"):
                fc.BarColor.Color = color_to_ole(fmt["fill_color"])
            results.append(f"{params['range']} 数据条已添加")

        elif cf_type == "color_scale":
            fc = rng.FormatConditions.AddColorScale(ColorScaleType=params.get("scale_count", 3))
            results.append(f"{params['range']} 色阶已添加")

        elif cf_type == "icon_set":
            fc = rng.FormatConditions.AddIconSetCondition()
            results.append(f"{params['range']} 图标集已添加")

        elif cf_type == "duplicate":
            fc = rng.FormatConditions.AddUniqueValues()
            fc.DupeUnique = 1  # xlDuplicate
            _apply_cf_format(fc, fmt)
            results.append(f"{params['range']} 重复值高亮已添加")

        elif cf_type == "unique":
            fc = rng.FormatConditions.AddUniqueValues()
            fc.DupeUnique = 0  # xlUnique
            _apply_cf_format(fc, fmt)
            results.append(f"{params['range']} 唯一值高亮已添加")

        elif cf_type == "clear":
            rng.FormatConditions.Delete()
            results.append(f"{params['range']} 条件格式已清除")

    # ===== 冻结窗格 =====
    elif action == "freeze_panes":
        cell = params.get("cell", "A2")
        sheet.Activate()
        excel.ActiveWindow.FreezePanes = False
        sheet.Range(cell).Select()
        excel.ActiveWindow.FreezePanes = True
        results.append(f"冻结窗格于 {cell}")

    elif action == "unfreeze_panes":
        sheet.Activate()
        excel.ActiveWindow.FreezePanes = False
        results.append("已取消冻结窗格")

    # ===== 隐藏/显示行列 =====
    elif action == "hide_rows":
        rows = params["rows"]  # "5:10" 或 5
        sheet.Rows(rows).Hidden = True
        results.append(f"行 {rows} 已隐藏")

    elif action == "show_rows":
        rows = params["rows"]
        sheet.Rows(rows).Hidden = False
        results.append(f"行 {rows} 已显示")

    elif action == "hide_cols":
        cols = params["cols"]  # "C:E" 或 "C"
        sheet.Columns(cols).Hidden = True
        results.append(f"列 {cols} 已隐藏")

    elif action == "show_cols":
        cols = params["cols"]
        sheet.Columns(cols).Hidden = False
        results.append(f"列 {cols} 已显示")

    # ===== 数据验证（下拉列表） =====
    elif action == "data_validation":
        rng = sheet.Range(params["range"])
        dv_type = params.get("dv_type", "list")
        if dv_type == "list":
            items = params.get("items") or params.get("formula1")
            if items is None:
                return {"error": "data_validation list 类型需要 items 或 formula1 参数"}
            if isinstance(items, (list, tuple)):
                items = ",".join(str(x) for x in items)
            try:
                rng.Validation.Delete()
            except Exception:
                pass
            # 关键修复：必须显式传 Operator，否则中文 Excel 会抛 COM 错误
            # xlValidateList=3, xlValidAlertStop=1, xlBetween=1
            rng.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=str(items))
            try:
                rng.Validation.IgnoreBlank = True
                rng.Validation.InCellDropdown = True
            except Exception:
                pass
            results.append(f"{params['range']} 下拉列表已设置")
        elif dv_type == "clear":
            rng.Validation.Delete()
            results.append(f"{params['range']} 数据验证已清除")

    # ===== 批注 =====
    elif action == "add_comment":
        cell = sheet.Range(params["cell"])
        text = params["text"]
        try:
            cell.Comment.Delete()
        except:
            pass
        cell.AddComment(text)
        results.append(f"{params['cell']} 批注已添加")

    elif action == "delete_comment":
        cell = sheet.Range(params["cell"])
        try:
            cell.Comment.Delete()
            results.append(f"{params['cell']} 批注已删除")
        except:
            results.append(f"{params['cell']} 没有批注")

    # ===== 超链接 =====
    elif action == "add_hyperlink":
        cell = sheet.Range(params["cell"])
        url = params["url"]
        text = params.get("text", url)
        sheet.Hyperlinks.Add(Anchor=cell, Address=url, TextToDisplay=text)
        results.append(f"{params['cell']} 超链接已添加: {url}")

    # ===== 清空 =====
    elif action == "clear_contents":
        rng = sheet.Range(params["range"])
        rng.ClearContents()
        results.append(f"{params['range']} 内容已清空")

    elif action == "clear_all":
        rng = sheet.Range(params["range"])
        rng.Clear()
        results.append(f"{params['range']} 已全部清空（值+样式）")

    # ===== 复制粘贴 =====
    elif action == "copy_paste":
        src = sheet.Range(params["source"])
        dst = sheet.Range(params["destination"])
        src.Copy(dst)
        results.append(f"{params['source']} → {params['destination']} 复制完成")

    # ===== 创建图表 =====
    elif action == "create_chart":
        # 图表类型映射 (xlChartType)
        chart_type_map = {
            "column": 51,          # xlColumnClustered
            "column_clustered": 51,
            "column_stacked": 52,  # xlColumnStacked
            "bar": 57,             # xlBarClustered
            "bar_clustered": 57,
            "bar_stacked": 58,     # xlBarStacked
            "line": 4,             # xlLine
            "line_markers": 65,    # xlLineMarkers
            "pie": 5,              # xlPie
            "pie_3d": -4102,       # xl3DPie
            "doughnut": -4120,     # xlDoughnut
            "scatter": -4169,      # xlXYScatter
            "scatter_lines": 74,   # xlXYScatterLines
            "area": 1,             # xlArea
            "area_stacked": 76,    # xlAreaStacked
            "radar": -4151,        # xlRadar
            "combo": -1,           # 特殊处理
            "柱状图": 51, "条形图": 57, "折线图": 4, "饼图": 5,
            "散点图": -4169, "面积图": 1, "环形图": -4120, "雷达图": -4151,
        }
        data_range = params.get("range")
        if not data_range:
            return {"error": "create_chart 需要 range 参数（数据区域，如 'A1:D10'）"}
        ct = params.get("chart_type", "column")
        ct_val = chart_type_map.get(str(ct).lower(), ct)
        if isinstance(ct_val, str):
            try:
                ct_val = int(ct_val)
            except ValueError:
                ct_val = 51  # 默认柱状图

        rng = sheet.Range(data_range)

        # === position 解析：支持单元格 "A7" 或单元格区域 "A7:F22" ===
        # - "A7"     → 仅定位左上角，width/height 用参数或默认
        # - "A7:F22" → 自动覆盖整个区域（Left/Top/Width/Height 都按区域算）
        pos = params.get("position")
        chart_left = rng.Left
        chart_top = rng.Top + rng.Height + 10
        chart_width = params.get("width", 480)
        chart_height = params.get("height", 300)
        position_used_range = False
        if pos:
            try:
                pos_rng = sheet.Range(str(pos))
                chart_left = pos_rng.Left
                chart_top = pos_rng.Top
                # 如果是区域（多个单元格），用区域的 Width/Height 作为图表尺寸
                if pos_rng.Cells.Count > 1:
                    chart_width = pos_rng.Width
                    chart_height = pos_rng.Height
                    position_used_range = True
            except Exception:
                pass  # position 解析失败时回退到默认位置

        chart_obj = sheet.Shapes.AddChart2(-1, ct_val, chart_left, chart_top, chart_width, chart_height)
        chart = chart_obj.Chart
        chart.SetSourceData(rng)

        # === 标题（显式控制 HasTitle）===
        # title=None  → 不动
        # title=""    → 显式不显示标题（HasTitle=False）
        # title="..." → 显示该标题
        if "title" in params:
            title = params.get("title")
            if title is None or title == "":
                chart.HasTitle = False
            else:
                chart.HasTitle = True
                chart.ChartTitle.Text = str(title)

        # === 图例 ===
        if params.get("show_legend") is False:
            chart.HasLegend = False
        elif params.get("show_legend") is True:
            chart.HasLegend = True

        # === 坐标轴标题 / 范围 ===
        # Excel COM: chart.Axes(xlCategory=1, xlPrimary=1)  →  X 轴
        #            chart.Axes(xlValue=2, xlPrimary=1)     →  Y 轴
        def _set_axis(axis_obj, title_text, min_val, max_val):
            if title_text is not None:
                if title_text == "":
                    axis_obj.HasTitle = False
                else:
                    axis_obj.HasTitle = True
                    axis_obj.AxisTitle.Text = str(title_text)
            if min_val is not None:
                try: axis_obj.MinimumScale = float(min_val)
                except Exception: pass
            if max_val is not None:
                try: axis_obj.MaximumScale = float(max_val)
                except Exception: pass

        try:
            x_title = params.get("x_axis_title")
            x_min = params.get("x_axis_min")
            x_max = params.get("x_axis_max")
            if x_title is not None or x_min is not None or x_max is not None:
                _set_axis(chart.Axes(1, 1), x_title, x_min, x_max)
        except Exception:
            pass
        try:
            y_title = params.get("y_axis_title")
            y_min = params.get("y_axis_min")
            y_max = params.get("y_axis_max")
            if y_title is not None or y_min is not None or y_max is not None:
                _set_axis(chart.Axes(2, 1), y_title, y_min, y_max)
        except Exception:
            pass

        # === 数据标签 ===
        # show_data_labels: True / False / None(不动)
        # data_labels: 同上别名
        sdl = params.get("show_data_labels")
        if sdl is None:
            sdl = params.get("data_labels")
        if sdl is True:
            try:
                for i in range(1, chart.SeriesCollection().Count + 1):
                    s = chart.SeriesCollection(i)
                    s.HasDataLabels = True
                    try: s.DataLabels().ShowValue = True
                    except Exception: pass
            except Exception:
                pass
        elif sdl is False:
            try:
                for i in range(1, chart.SeriesCollection().Count + 1):
                    chart.SeriesCollection(i).HasDataLabels = False
            except Exception:
                pass

        # === 兼容旧的 left/top 直接覆盖（优先级最高）===
        if "left" in params:
            chart_obj.Left = params["left"]
        if "top" in params:
            chart_obj.Top = params["top"]

        pos_desc = f"，位置: {pos}" + ("[区域]" if position_used_range else "[左上角]") if pos else ""
        results.append(f"图表已创建（类型: {ct}，数据: {data_range}{pos_desc}）")

    # ===== 数字格式 =====
    elif action == "number_format":
        # 常用格式别名
        format_alias = {
            "number": "0.00", "integer": "0", "percent": "0.00%",
            "currency": "¥#,##0.00", "currency_usd": "$#,##0.00",
            "date": "yyyy-mm-dd", "datetime": "yyyy-mm-dd hh:mm:ss",
            "time": "hh:mm:ss", "text": "@", "scientific": "0.00E+00",
            "thousands": "#,##0", "thousands_decimal": "#,##0.00",
            "百分比": "0.00%", "货币": "¥#,##0.00", "日期": "yyyy-mm-dd",
            "文本": "@", "科学计数": "0.00E+00",
        }
        operations = params.get("operations")
        if not operations:
            r = params.get("range") or params.get("cell")
            f = params.get("format_string") or params.get("format")
            if r and f:
                operations = [{"range": r, "format": f}]
            else:
                return {"error": "number_format 需要 operations 或 range+format_string 参数"}
        if isinstance(operations, dict):
            operations = [operations]
        for op in operations:
            rng = sheet.Range(op.get("range") or op.get("cell"))
            fmt = op.get("format") or op.get("format_string", "General")
            fmt = format_alias.get(str(fmt).lower(), fmt)
            rng.NumberFormat = fmt
            results.append(f"{op.get('range', op.get('cell'))} 格式已设置为 {fmt}")

    # ===== 查找替换 =====
    elif action == "find_replace":
        find_text = params.get("find")
        replace_text = params.get("replace", "")
        if not find_text:
            return {"error": "find_replace 需要 find 参数"}
        search_range = sheet.Range(params["range"]) if params.get("range") else sheet.UsedRange
        replaced = search_range.Replace(What=find_text, Replacement=replace_text,
                                        LookAt=2,  # xlPart
                                        MatchCase=params.get("match_case", False))
        if replaced:
            results.append(f"查找替换完成: '{find_text}' → '{replace_text}'")
        else:
            results.append(f"⚠️ 未找到 '{find_text}'")

    # ===== 插入图片 =====
    elif action == "insert_image":
        image_source = params.get("path") or params.get("image_path")
        if not image_source:
            return {"error": "insert_image 需要 path 参数"}
        image_path, temp_image_path, image_error = _prepare_office_image_path(image_source)
        if image_error:
            return {"error": image_error}
        cell_ref = params.get("cell", "A1")
        cell = sheet.Range(cell_ref)
        try:
            pic = sheet.Shapes.AddPicture(
                Filename=image_path,
                LinkToFile=False, SaveWithDocument=True,
                Left=cell.Left, Top=cell.Top,
                Width=params.get("width", -1), Height=params.get("height", -1)
            )
            # -1 = 保持原始尺寸
            if params.get("width", -1) == -1:
                pic.ScaleWidth(1, True)
            if params.get("height", -1) == -1:
                pic.ScaleHeight(1, True)
            results.append(f"图片已插入到 {cell_ref}: {image_source}")
        finally:
            _safe_remove(temp_image_path)

    # ===== 自动调整列宽/行高 =====
    elif action == "auto_fit":
        target = params.get("range") or params.get("cols") or params.get("rows")
        if not target:
            # 自动调整整个使用区域
            sheet.UsedRange.Columns.AutoFit()
            sheet.UsedRange.Rows.AutoFit()
            results.append("已自动调整所有列宽和行高")
        else:
            rng = sheet.Range(target) if params.get("range") else None
            fit_type = params.get("fit_type", "both")  # "columns", "rows", "both"
            if params.get("cols"):
                sheet.Columns(params["cols"]).AutoFit()
                results.append(f"列 {params['cols']} 已自动调整宽度")
            elif params.get("rows"):
                sheet.Rows(params["rows"]).AutoFit()
                results.append(f"行 {params['rows']} 已自动调整高度")
            elif rng:
                if fit_type in ("columns", "both"):
                    rng.Columns.AutoFit()
                if fit_type in ("rows", "both"):
                    rng.Rows.AutoFit()
                results.append(f"{target} 已自动调整{'列宽和行高' if fit_type == 'both' else '列宽' if fit_type == 'columns' else '行高'}")

    # ===== 工作表保护 =====
    elif action == "protect":
        password = params.get("password", "")
        allow_formatting = params.get("allow_formatting", False)
        if password:
            sheet.Protect(Password=password, AllowFormattingCells=allow_formatting)
        else:
            sheet.Protect(AllowFormattingCells=allow_formatting)
        results.append(f"工作表 '{sheet.Name}' 已保护")

    elif action == "unprotect":
        password = params.get("password", "")
        if password:
            sheet.Unprotect(Password=password)
        else:
            sheet.Unprotect()
        results.append(f"工作表 '{sheet.Name}' 已取消保护")

    # ===== 命名范围 =====
    elif action == "named_range":
        nr_action = params.get("nr_action", "create")
        if nr_action == "create":
            name = params.get("name")
            ref_range = params.get("range")
            if not name or not ref_range:
                return {"error": "named_range create 需要 name 和 range 参数"}
            wb.Names.Add(Name=name, RefersTo=f"={sheet.Name}!{ref_range}")
            results.append(f"命名范围 '{name}' → {ref_range} 已创建")
        elif nr_action == "delete":
            name = params.get("name")
            if not name:
                return {"error": "named_range delete 需要 name 参数"}
            try:
                wb.Names(name).Delete()
                results.append(f"命名范围 '{name}' 已删除")
            except Exception as e:
                return {"error": f"删除命名范围失败: {e}"}
        elif nr_action == "list":
            names_list = []
            for i in range(1, wb.Names.Count + 1):
                n = wb.Names(i)
                names_list.append({"name": n.Name, "refers_to": n.RefersTo, "visible": n.Visible})
            results.append(f"共 {len(names_list)} 个命名范围")
            return {"success": True, "results": results, "named_ranges": names_list}
        else:
            return {"error": f"未知的 nr_action: {nr_action}，支持 create/delete/list"}

    # ===== 数据透视表 =====
    elif action == "pivot_table":
        source_range = params.get("range")
        if not source_range:
            return {"error": "pivot_table 需要 range 参数（源数据区域）"}
        dest_cell = params.get("destination", "")
        pt_name = params.get("name", f"PivotTable_{len(results)+1}")

        src_rng = sheet.Range(source_range)
        # 确定目标位置
        if dest_cell:
            # 如果指定了目标工作表
            dest_sheet_name = params.get("dest_sheet")
            if dest_sheet_name:
                try:
                    dest_sheet = wb.Sheets(dest_sheet_name)
                except Exception:
                    dest_sheet = wb.Sheets.Add()
                    dest_sheet.Name = dest_sheet_name
                dest_rng = dest_sheet.Range(dest_cell)
            else:
                dest_rng = sheet.Range(dest_cell)
        else:
            # 创建新工作表放透视表
            new_sheet = wb.Sheets.Add()
            new_sheet.Name = pt_name[:31]  # 工作表名最长31字符
            dest_rng = new_sheet.Range("A3")

        # 创建透视表缓存
        pc = wb.PivotCaches().Create(
            SourceType=1,  # xlDatabase
            SourceData=src_rng
        )
        # 创建透视表
        pt = pc.CreatePivotTable(
            TableDestination=dest_rng,
            TableName=pt_name
        )

        # 添加行字段
        row_fields = params.get("row_fields", [])
        for field_name in row_fields:
            try:
                pf = pt.PivotFields(field_name)
                pf.Orientation = 1  # xlRowField
            except Exception as e:
                results.append(f"⚠️ 行字段 '{field_name}' 添加失败: {e}")

        # 添加列字段
        col_fields = params.get("col_fields", [])
        for field_name in col_fields:
            try:
                pf = pt.PivotFields(field_name)
                pf.Orientation = 2  # xlColumnField
            except Exception as e:
                results.append(f"⚠️ 列字段 '{field_name}' 添加失败: {e}")

        # 添加值字段
        value_fields = params.get("value_fields", [])
        for vf in value_fields:
            if isinstance(vf, str):
                vf = {"field": vf, "function": "sum"}
            field_name = vf.get("field")
            func_name = vf.get("function", "sum").lower()
            func_map = {
                "sum": -4157, "count": -4112, "average": -4106,
                "max": -4136, "min": -4139, "product": -4149,
            }
            try:
                pf = pt.PivotFields(field_name)
                pf.Orientation = 4  # xlDataField
                pf.Function = func_map.get(func_name, -4157)
            except Exception as e:
                results.append(f"⚠️ 值字段 '{field_name}' 添加失败: {e}")

        # 添加筛选字段
        filter_fields = params.get("filter_fields", [])
        for field_name in filter_fields:
            try:
                pf = pt.PivotFields(field_name)
                pf.Orientation = 3  # xlPageField
            except Exception as e:
                results.append(f"⚠️ 筛选字段 '{field_name}' 添加失败: {e}")

        results.append(f"数据透视表 '{pt_name}' 已创建（源: {source_range}）")

    # ============================================================
    # 批次 A：图表编辑/美化
    # ============================================================

    # ===== 列出当前工作表所有图表 =====
    elif action == "list_charts":
        charts_info = []
        include_series = params.get("include_series", True)
        try:
            shapes = sheet.Shapes
            chart_idx = 0
            for i in range(1, shapes.Count + 1):
                sh = shapes(i)
                if sh.HasChart:
                    chart_idx += 1
                    brief = _read_chart_brief(sh, include_series=include_series)
                    brief["chart_index"] = chart_idx
                    charts_info.append(brief)
        except Exception as exc:
            return {"error": f"列出图表失败: {exc}"}
        return {"success": True, "sheet": sheet.Name, "chart_count": len(charts_info), "charts": charts_info}

    # ===== 编辑已有图表 =====
    elif action == "edit_chart":
        identifier = params.get("chart_index") or params.get("chart_name") or params.get("chart_title")
        if identifier is None:
            return {"error": "edit_chart 需要 chart_index（1-based）/chart_name/chart_title 之一定位图表"}
        chart_shape = _find_chart_by_identifier(sheet, identifier)
        if chart_shape is None:
            return {"error": f"未找到图表: {identifier}（请用 list_charts 查看可用图表）"}
        chart = chart_shape.Chart

        # 改类型
        if params.get("chart_type") is not None:
            try:
                chart.ChartType = _resolve_chart_type(params["chart_type"])
                results.append(f"图表类型 → {params['chart_type']}")
            except Exception as exc:
                errors.append(f"修改图表类型失败: {exc}")

        # 改数据源
        if params.get("range"):
            try:
                chart.SetSourceData(sheet.Range(params["range"]))
                results.append(f"数据源 → {params['range']}")
            except Exception as exc:
                errors.append(f"修改数据源失败: {exc}")

        # 位置/大小
        if params.get("position"):
            try:
                pos_rng = sheet.Range(str(params["position"]))
                chart_shape.Left = pos_rng.Left
                chart_shape.Top = pos_rng.Top
                if pos_rng.Cells.Count > 1:
                    chart_shape.Width = pos_rng.Width
                    chart_shape.Height = pos_rng.Height
                results.append(f"位置 → {params['position']}")
            except Exception as exc:
                errors.append(f"修改位置失败: {exc}")
        if params.get("left") is not None:
            chart_shape.Left = float(params["left"])
            results.append(f"left → {params['left']}")
        if params.get("top") is not None:
            chart_shape.Top = float(params["top"])
            results.append(f"top → {params['top']}")
        if params.get("width") is not None:
            chart_shape.Width = float(params["width"])
            results.append(f"width → {params['width']}")
        if params.get("height") is not None:
            chart_shape.Height = float(params["height"])
            results.append(f"height → {params['height']}")

        # 标题
        if "title" in params:
            title = params["title"]
            if title is None or title == "":
                chart.HasTitle = False
                results.append("标题 → 已隐藏")
            else:
                chart.HasTitle = True
                chart.ChartTitle.Text = str(title)
                results.append(f"标题 → {title}")
        if isinstance(params.get("title_font"), dict):
            try:
                if not chart.HasTitle:
                    chart.HasTitle = True
                _apply_text_font(chart.ChartTitle.Format.TextFrame2.TextRange.Font, params["title_font"])
                results.append("标题字体已设置")
            except Exception as exc:
                errors.append(f"标题字体设置失败: {exc}")

        # 图例
        if params.get("show_legend") is False:
            chart.HasLegend = False
            results.append("图例 → 隐藏")
        elif params.get("show_legend") is True:
            chart.HasLegend = True
            results.append("图例 → 显示")
        if params.get("legend_position") is not None and chart.HasLegend:
            try:
                lp = params["legend_position"]
                pos_val = LEGEND_POSITION_MAP.get(str(lp).lower(), lp)
                chart.Legend.Position = int(pos_val)
                results.append(f"图例位置 → {lp}")
            except Exception as exc:
                errors.append(f"图例位置设置失败: {exc}")
        if isinstance(params.get("legend_font"), dict) and chart.HasLegend:
            try:
                _apply_text_font(chart.Legend.Format.TextFrame2.TextRange.Font, params["legend_font"])
                results.append("图例字体已设置")
            except Exception as exc:
                errors.append(f"图例字体设置失败: {exc}")

        # 坐标轴
        x_axis_cfg = params.get("x_axis") or {}
        # 旧风格兼容
        for legacy_key, axis_key in [
            ("x_axis_title", "title"), ("x_axis_min", "min"), ("x_axis_max", "max"),
        ]:
            if params.get(legacy_key) is not None:
                x_axis_cfg.setdefault(axis_key, params[legacy_key])
        y_axis_cfg = params.get("y_axis") or {}
        for legacy_key, axis_key in [
            ("y_axis_title", "title"), ("y_axis_min", "min"), ("y_axis_max", "max"),
        ]:
            if params.get(legacy_key) is not None:
                y_axis_cfg.setdefault(axis_key, params[legacy_key])
        if x_axis_cfg:
            try:
                _apply_chart_axis_config(chart.Axes(1, 1), x_axis_cfg)
                results.append("X 轴已配置")
            except Exception as exc:
                errors.append(f"X 轴配置失败: {exc}")
        if y_axis_cfg:
            try:
                _apply_chart_axis_config(chart.Axes(2, 1), y_axis_cfg)
                results.append("Y 轴已配置")
            except Exception as exc:
                errors.append(f"Y 轴配置失败: {exc}")

        # 网格线
        gl = params.get("gridlines")
        if isinstance(gl, dict):
            try:
                y_axis = chart.Axes(2, 1)
                if "major_y" in gl:
                    y_axis.HasMajorGridlines = bool(gl["major_y"])
                if "minor_y" in gl:
                    y_axis.HasMinorGridlines = bool(gl["minor_y"])
                if isinstance(gl.get("major_y_color"), (str, list, tuple)) and y_axis.HasMajorGridlines:
                    y_axis.MajorGridlines.Format.Line.ForeColor.RGB = color_to_ole(gl["major_y_color"])
            except Exception:
                pass
            try:
                x_axis = chart.Axes(1, 1)
                if "major_x" in gl:
                    x_axis.HasMajorGridlines = bool(gl["major_x"])
                if "minor_x" in gl:
                    x_axis.HasMinorGridlines = bool(gl["minor_x"])
            except Exception:
                pass
            results.append("网格线已设置")

        # 绘图区/图表区填充
        pa = params.get("plot_area")
        if isinstance(pa, dict):
            try:
                if pa.get("fill_color") is not None:
                    chart.PlotArea.Format.Fill.Visible = -1
                    chart.PlotArea.Format.Fill.Solid()
                    chart.PlotArea.Format.Fill.ForeColor.RGB = color_to_ole(pa["fill_color"])
                if pa.get("border_color") is not None:
                    chart.PlotArea.Format.Line.Visible = -1
                    chart.PlotArea.Format.Line.ForeColor.RGB = color_to_ole(pa["border_color"])
                results.append("绘图区已设置")
            except Exception as exc:
                errors.append(f"绘图区设置失败: {exc}")
        ca = params.get("chart_area")
        if isinstance(ca, dict):
            try:
                if ca.get("fill_color") is not None:
                    chart.ChartArea.Format.Fill.Visible = -1
                    chart.ChartArea.Format.Fill.Solid()
                    chart.ChartArea.Format.Fill.ForeColor.RGB = color_to_ole(ca["fill_color"])
                if ca.get("border_color") is not None:
                    chart.ChartArea.Format.Line.Visible = -1
                    chart.ChartArea.Format.Line.ForeColor.RGB = color_to_ole(ca["border_color"])
                if isinstance(ca.get("font"), dict):
                    _apply_text_font(chart.ChartArea.Format.TextFrame2.TextRange.Font, ca["font"])
                results.append("图表区已设置")
            except Exception as exc:
                errors.append(f"图表区设置失败: {exc}")

        # 系列格式（series_format = [{index:1, fill_color:"红", ...}]）
        sf_list = params.get("series_format")
        if isinstance(sf_list, dict):
            sf_list = [sf_list]
        if isinstance(sf_list, list):
            for s_fmt in sf_list:
                if not isinstance(s_fmt, dict):
                    continue
                idx = s_fmt.get("index", 1)
                try:
                    s_obj = chart.SeriesCollection(int(idx))
                    _apply_series_format(s_obj, s_fmt)
                    results.append(f"系列 {idx} 格式已设置")
                except Exception as exc:
                    errors.append(f"系列 {idx} 格式设置失败: {exc}")

        # 数据标签（全局开关 + 格式）
        sdl = params.get("show_data_labels")
        if sdl is None:
            sdl = params.get("data_labels")
        if sdl is not None:
            try:
                sc = chart.SeriesCollection()
                for i in range(1, sc.Count + 1):
                    s = sc(i)
                    s.HasDataLabels = bool(sdl)
                    if sdl:
                        try: s.DataLabels().ShowValue = True
                        except Exception: pass
                results.append(f"数据标签 → {'显示' if sdl else '隐藏'}")
            except Exception as exc:
                errors.append(f"数据标签设置失败: {exc}")

        # 趋势线（trendline = [{series_index, type, name, forward, backward, show_equation, show_r_squared}]）
        tl_list = params.get("trendlines")
        if isinstance(tl_list, dict):
            tl_list = [tl_list]
        if isinstance(tl_list, list):
            trendline_type_map = {
                "linear": -4132, "线性": -4132,
                "logarithmic": -4133, "对数": -4133,
                "polynomial": 3, "多项式": 3,
                "power": 4, "幂": 4,
                "exponential": 5, "指数": 5,
                "moving_average": 6, "移动平均": 6, "moving_avg": 6,
            }
            for tl in tl_list:
                if not isinstance(tl, dict):
                    continue
                s_idx = tl.get("series_index", 1)
                try:
                    s_obj = chart.SeriesCollection(int(s_idx))
                    tl_type = tl.get("type", "linear")
                    tl_type_val = trendline_type_map.get(str(tl_type).lower(), tl_type)
                    new_tl = s_obj.Trendlines().Add(Type=int(tl_type_val))
                    if tl.get("name"):
                        new_tl.Name = str(tl["name"])
                    if tl.get("forward") is not None:
                        new_tl.Forward = float(tl["forward"])
                    if tl.get("backward") is not None:
                        new_tl.Backward = float(tl["backward"])
                    if tl.get("order") is not None and int(tl_type_val) == 3:
                        new_tl.Order = int(tl["order"])
                    if tl.get("period") is not None and int(tl_type_val) == 6:
                        new_tl.Period = int(tl["period"])
                    if tl.get("show_equation"):
                        new_tl.DisplayEquation = True
                    if tl.get("show_r_squared"):
                        new_tl.DisplayRSquared = True
                    results.append(f"系列 {s_idx} 添加趋势线 {tl_type}")
                except Exception as exc:
                    errors.append(f"添加趋势线失败: {exc}")

        # 应用内置 ChartStyle（1-48）
        if params.get("chart_style") is not None:
            try:
                chart.ChartStyle = int(params["chart_style"])
                results.append(f"内置样式 → {params['chart_style']}")
            except Exception as exc:
                errors.append(f"应用 ChartStyle 失败: {exc}")

    # ===== 删除图表 =====
    elif action == "delete_chart":
        identifier = params.get("chart_index") or params.get("chart_name") or params.get("chart_title")
        if identifier is None:
            return {"error": "delete_chart 需要 chart_index/chart_name/chart_title 之一"}
        chart_shape = _find_chart_by_identifier(sheet, identifier)
        if chart_shape is None:
            return {"error": f"未找到图表: {identifier}"}
        chart_name = chart_shape.Name
        chart_shape.Delete()
        results.append(f"图表 '{chart_name}' 已删除")

    # ===== 套用内置图表样式快捷入口 =====
    elif action == "chart_style":
        identifier = params.get("chart_index") or params.get("chart_name") or params.get("chart_title")
        if identifier is None:
            return {"error": "chart_style 需要 chart_index/chart_name/chart_title 之一"}
        chart_shape = _find_chart_by_identifier(sheet, identifier)
        if chart_shape is None:
            return {"error": f"未找到图表: {identifier}"}
        style_id = params.get("style") or params.get("style_id") or params.get("chart_style")
        if style_id is None:
            return {"error": "chart_style 需要 style 参数（1-48 的整数）"}
        try:
            chart_shape.Chart.ChartStyle = int(style_id)
            results.append(f"图表样式 → {style_id}")
        except Exception as exc:
            return {"error": f"应用 ChartStyle 失败: {exc}"}

    # ===== 导出图表为 PNG =====
    elif action == "export_chart":
        identifier = params.get("chart_index") or params.get("chart_name") or params.get("chart_title")
        if identifier is None:
            return {"error": "export_chart 需要 chart_index/chart_name/chart_title 之一"}
        chart_shape = _find_chart_by_identifier(sheet, identifier)
        if chart_shape is None:
            return {"error": f"未找到图表: {identifier}"}
        export_path = params.get("path") or params.get("output_path")
        if not export_path:
            export_path = os.path.join(tempfile.gettempdir(), f"chart_{uuid.uuid4().hex[:8]}.png")
        else:
            os.makedirs(os.path.dirname(export_path) or ".", exist_ok=True)
        export_format = (params.get("format") or "PNG").upper()
        try:
            chart_shape.Chart.Export(Filename=export_path, FilterName=export_format)
        except Exception as exc:
            return {"error": f"导出图表失败: {exc}"}
        result_info = {"success": True, "results": [f"图表已导出 → {export_path}"], "path": export_path}
        # 可选回传内嵌图片（默认不开，避免 base64 爆上下文）
        if params.get("inline_image"):
            try:
                with open(export_path, "rb") as f:
                    blob = f.read()
                entry = _build_image_entry(os.path.basename(export_path), blob, "image/png")
                result_info["images"] = [entry]
                result_info["image_count"] = 1
            except Exception:
                pass
        return result_info

    # ============================================================
    # 批次 B：单元格样式扩展
    # ============================================================

    # ===== 套用 Excel 内置 CellStyle =====
    elif action == "apply_cell_style":
        target = params.get("range") or params.get("cell")
        if not target:
            return {"error": "apply_cell_style 需要 range 或 cell 参数"}
        style_input = params.get("style") or params.get("style_name") or params.get("cell_style")
        if not style_input:
            return {"error": "apply_cell_style 需要 style 参数（如 'good'/'bad'/'heading_1'/'accent_1'/'20%_accent_1' 等）"}
        # 翻译
        style_name = EXCEL_BUILTIN_CELL_STYLES.get(str(style_input).strip().lower(),
                                                    EXCEL_BUILTIN_CELL_STYLES.get(str(style_input).strip(), str(style_input)))
        try:
            sheet.Range(target).Style = style_name
            results.append(f"{target} 已套用样式 '{style_name}'")
        except Exception as exc:
            return {"error": f"套用 CellStyle '{style_name}' 失败: {exc}"}

    # ===== 列出工作簿可用的 CellStyle 名 =====
    elif action == "list_cell_styles":
        names = []
        try:
            for i in range(1, wb.Styles.Count + 1):
                try:
                    names.append(wb.Styles(i).Name)
                except Exception:
                    pass
        except Exception as exc:
            return {"error": f"枚举 CellStyle 失败: {exc}"}
        return {"success": True, "count": len(names), "styles": names}

    # ===== 渐变填充 =====
    elif action == "gradient_fill":
        target = params.get("range") or params.get("cell")
        if not target:
            return {"error": "gradient_fill 需要 range 或 cell 参数"}
        color1 = params.get("color1") or params.get("from_color")
        color2 = params.get("color2") or params.get("to_color")
        if color1 is None or color2 is None:
            return {"error": "gradient_fill 需要 color1 和 color2"}
        direction = params.get("direction", "horizontal")
        direction_val = GRADIENT_STYLE_MAP.get(str(direction).lower(), direction)
        variant = int(params.get("variant", 1))  # 1-4，渐变变体
        try:
            rng = sheet.Range(target)
            rng.Interior.Pattern = -4124  # xlPatternLightGray, 触发 Gradient 初始化
            # Excel Range.Interior 不直接支持渐变；用 ColorIndex/Pattern 替代
            # 真正渐变只能用 Shape 的 Fill；对 Range 用单色或图案
            # 因此用 PatternColorIndex + PatternColor 模拟两色图案
            rng.Interior.Color = color_to_ole(color1)
            rng.Interior.PatternColor = color_to_ole(color2)
            try:
                rng.Interior.Pattern = -4128  # 默认水平线条
                if str(direction).lower() in ("vertical", "垂直"):
                    rng.Interior.Pattern = -4166
                elif str(direction).lower() in ("diagonal_down", "对角向下", "右斜线"):
                    rng.Interior.Pattern = -4121
                elif str(direction).lower() in ("diagonal_up", "对角向上", "左斜线"):
                    rng.Interior.Pattern = -4162
            except Exception:
                pass
            results.append(f"{target} 已应用渐变填充（{direction}）")
            results.append("⚠️ Excel Range 不支持真正的两色渐变，已用 Pattern + PatternColor 近似（如需真实渐变请插入 Shape 覆盖）")
        except Exception as exc:
            return {"error": f"渐变填充失败: {exc}"}

    # ===== 图案填充 =====
    elif action == "pattern_fill":
        target = params.get("range") or params.get("cell")
        if not target:
            return {"error": "pattern_fill 需要 range 或 cell 参数"}
        pattern = params.get("pattern", "solid")
        pattern_val = _resolve_pattern_fill(pattern)
        if pattern_val is None:
            return {"error": f"无法解析 pattern: {pattern}，可用: {list(PATTERN_FILL_MAP.keys())}"}
        try:
            rng = sheet.Range(target)
            if params.get("fill_color") is not None:
                rng.Interior.Color = color_to_ole(params["fill_color"])
            rng.Interior.Pattern = pattern_val
            if params.get("pattern_color") is not None:
                rng.Interior.PatternColor = color_to_ole(params["pattern_color"])
            results.append(f"{target} 已应用图案填充（{pattern}）")
        except Exception as exc:
            return {"error": f"图案填充失败: {exc}"}

    # ===== TintAndShade 调浅/调深 =====
    elif action == "tint_and_shade":
        target = params.get("range") or params.get("cell")
        if not target:
            return {"error": "tint_and_shade 需要 range 或 cell 参数"}
        tint = params.get("tint", params.get("value", 0))
        try:
            tint = max(-1.0, min(1.0, float(tint)))
        except Exception:
            return {"error": "tint 需为 -1.0 ~ 1.0 之间的浮点数（正数变浅、负数变深）"}
        try:
            rng = sheet.Range(target)
            apply_to = params.get("apply_to", "fill")  # fill / font / both
            if apply_to in ("fill", "both"):
                rng.Interior.TintAndShade = tint
            if apply_to in ("font", "both"):
                rng.Font.TintAndShade = tint
            results.append(f"{target} TintAndShade={tint}（{apply_to}）")
        except Exception as exc:
            return {"error": f"应用 TintAndShade 失败: {exc}"}

    # ===== 斜线表头 =====
    elif action == "diagonal_borders":
        target = params.get("range") or params.get("cell")
        if not target:
            return {"error": "diagonal_borders 需要 range 或 cell 参数"}
        direction = params.get("direction", "down")  # down / up / both
        weight = params.get("weight", "thin")
        color = params.get("color", "黑")
        style = params.get("style", "continuous")
        try:
            rng = sheet.Range(target)
            border_style = {"style": style, "weight": weight, "color": color}
            if direction in ("down", "both", "对角向下", "下斜", "右下"):
                _apply_excel_border_style(rng.Borders(BORDER_INDEX_MAP["diagonal_down"]), border_style)
            if direction in ("up", "both", "对角向上", "上斜", "右上"):
                _apply_excel_border_style(rng.Borders(BORDER_INDEX_MAP["diagonal_up"]), border_style)
            results.append(f"{target} 已添加斜线（{direction}）")
        except Exception as exc:
            return {"error": f"斜线表头失败: {exc}"}

    # ===== 设置行高 =====
    elif action == "set_row_height":
        height = params.get("height")
        if height is None:
            return {"error": "set_row_height 需要 height 参数"}
        rows_target = params.get("rows") or params.get("row")
        try:
            if rows_target is None:
                sheet.UsedRange.EntireRow.RowHeight = float(height)
                results.append(f"全表行高 → {height}")
            else:
                sheet.Rows(rows_target).RowHeight = float(height)
                results.append(f"行 {rows_target} 高度 → {height}")
        except Exception as exc:
            return {"error": f"设置行高失败: {exc}"}

    # ===== 设置列宽 =====
    elif action == "set_column_width":
        width = params.get("width")
        if width is None:
            return {"error": "set_column_width 需要 width 参数"}
        cols_target = params.get("cols") or params.get("col") or params.get("columns")
        try:
            if cols_target is None:
                sheet.UsedRange.EntireColumn.ColumnWidth = float(width)
                results.append(f"全表列宽 → {width}")
            else:
                sheet.Columns(cols_target).ColumnWidth = float(width)
                results.append(f"列 {cols_target} 宽度 → {width}")
        except Exception as exc:
            return {"error": f"设置列宽失败: {exc}"}

    # ===== 格式刷（复制格式不带值）=====
    elif action == "copy_format":
        source = params.get("source")
        destination = params.get("destination")
        if not source or not destination:
            return {"error": "copy_format 需要 source 和 destination"}
        try:
            src_rng = sheet.Range(source)
            dst_rng = sheet.Range(destination)
            src_rng.Copy()
            # xlPasteFormats = -4122
            dst_rng.PasteSpecial(Paste=-4122)
            try:
                excel.CutCopyMode = False
            except Exception:
                pass
            results.append(f"格式 {source} → {destination}")
        except Exception as exc:
            return {"error": f"复制格式失败: {exc}"}

    # ============================================================
    # 批次 C：ListObject（Excel 表格 Ctrl+T）
    # ============================================================

    # ===== 列出 ListObject =====
    elif action == "list_tables":
        items = _read_excel_list_objects(sheet)
        return {"success": True, "sheet": sheet.Name, "count": len(items), "tables": items}

    # ===== 区域 → 表格（ListObject）=====
    elif action == "convert_to_table":
        source_range = params.get("range")
        if not source_range:
            return {"error": "convert_to_table 需要 range 参数"}
        has_headers = 1 if params.get("has_headers", True) else 2  # xlYes=1, xlNo=2
        table_name = params.get("name") or params.get("table_name")
        style_name = params.get("style") or params.get("style_name") or "TableStyleMedium2"
        # 别名快捷
        style_alias = {
            "light_1": "TableStyleLight1", "light_2": "TableStyleLight2", "light_9": "TableStyleLight9",
            "medium_1": "TableStyleMedium1", "medium_2": "TableStyleMedium2", "medium_9": "TableStyleMedium9",
            "dark_1": "TableStyleDark1", "dark_2": "TableStyleDark2",
            "浅色1": "TableStyleLight1", "中度2": "TableStyleMedium2", "深色1": "TableStyleDark1",
        }
        style_name = style_alias.get(str(style_name).lower(), style_name)
        try:
            rng = sheet.Range(source_range)
            lo = sheet.ListObjects.Add(SourceType=1,  # xlSrcRange
                                       Source=rng,
                                       XlListObjectHasHeaders=has_headers)
            if table_name:
                try:
                    lo.Name = str(table_name)
                except Exception:
                    pass
            try:
                lo.TableStyle = style_name
            except Exception:
                pass
            results.append(f"区域 {source_range} 已转为表格 '{lo.Name}'（样式 {style_name}）")
        except Exception as exc:
            return {"error": f"转换为表格失败: {exc}"}

    # ===== 修改表格样式 =====
    elif action == "table_style":
        table_id = params.get("table_index") or params.get("table_name") or params.get("name")
        if table_id is None:
            return {"error": "table_style 需要 table_index 或 table_name 定位"}
        try:
            lo = sheet.ListObjects(table_id if isinstance(table_id, str) else int(table_id))
        except Exception:
            return {"error": f"未找到表格: {table_id}"}
        style_name = params.get("style") or params.get("style_name")
        if style_name:
            style_alias = {
                "light_1": "TableStyleLight1", "light_2": "TableStyleLight2", "light_9": "TableStyleLight9",
                "medium_1": "TableStyleMedium1", "medium_2": "TableStyleMedium2", "medium_9": "TableStyleMedium9",
                "dark_1": "TableStyleDark1", "dark_2": "TableStyleDark2",
            }
            style_name = style_alias.get(str(style_name).lower(), style_name)
            try:
                lo.TableStyle = style_name
                results.append(f"表格 '{lo.Name}' 样式 → {style_name}")
            except Exception as exc:
                errors.append(f"应用表格样式失败: {exc}")
        for key, attr in [
            ("show_header", "ShowHeaders"),
            ("show_totals", "ShowTotals"),
            ("show_banded_rows", "ShowTableStyleRowStripes"),
            ("show_banded_columns", "ShowTableStyleColumnStripes"),
            ("show_first_column", "ShowTableStyleFirstColumn"),
            ("show_last_column", "ShowTableStyleLastColumn"),
        ]:
            if key in params:
                try:
                    setattr(lo, attr, bool(params[key]))
                    results.append(f"表格 '{lo.Name}' {key} → {params[key]}")
                except Exception as exc:
                    errors.append(f"设置 {key} 失败: {exc}")

    # ===== 表格汇总行 =====
    elif action == "table_total_row":
        table_id = params.get("table_index") or params.get("table_name") or params.get("name")
        if table_id is None:
            return {"error": "table_total_row 需要 table_index 或 table_name"}
        try:
            lo = sheet.ListObjects(table_id if isinstance(table_id, str) else int(table_id))
        except Exception:
            return {"error": f"未找到表格: {table_id}"}
        try:
            lo.ShowTotals = bool(params.get("show", True))
            results.append(f"表格 '{lo.Name}' 汇总行 → {'显示' if lo.ShowTotals else '隐藏'}")
        except Exception as exc:
            errors.append(f"切换汇总行失败: {exc}")
        # 为每列指定汇总函数：totals_calculations = [{column:"销售额", function:"sum"}]
        tc_list = params.get("totals_calculations") or params.get("totals")
        if isinstance(tc_list, list) and lo.ShowTotals:
            tc_map = {
                "none": 0, "无": 0,
                "average": 1, "平均": 1, "average_value": 1,
                "count": 2, "计数": 2,
                "count_nums": 3, "数值计数": 3,
                "max": 4, "最大": 4,
                "min": 5, "最小": 5,
                "sum": 6, "求和": 6, "合计": 6,
                "std_dev": 7, "标准差": 7,
                "var": 8, "方差": 8,
                "custom": 9, "自定义": 9,
            }
            for tc in tc_list:
                if not isinstance(tc, dict):
                    continue
                col_id = tc.get("column") or tc.get("col") or tc.get("name")
                func = tc.get("function") or tc.get("func", "sum")
                func_val = tc_map.get(str(func).lower(), func)
                try:
                    if isinstance(col_id, str):
                        col_obj = lo.ListColumns(col_id)
                    else:
                        col_obj = lo.ListColumns(int(col_id))
                    col_obj.TotalsCalculation = int(func_val)
                    results.append(f"表格 '{lo.Name}' 列 '{col_id}' 汇总 → {func}")
                except Exception as exc:
                    errors.append(f"列 '{col_id}' 汇总失败: {exc}")

    # ===== 重置表格范围 =====
    elif action == "table_resize":
        table_id = params.get("table_index") or params.get("table_name") or params.get("name")
        new_range = params.get("range") or params.get("new_range")
        if table_id is None or not new_range:
            return {"error": "table_resize 需要 table_index/table_name + range"}
        try:
            lo = sheet.ListObjects(table_id if isinstance(table_id, str) else int(table_id))
            lo.Resize(sheet.Range(new_range))
            results.append(f"表格 '{lo.Name}' 范围 → {new_range}")
        except Exception as exc:
            return {"error": f"重置表格范围失败: {exc}"}

    # ===== 表格转回普通区域 =====
    elif action == "convert_to_range":
        table_id = params.get("table_index") or params.get("table_name") or params.get("name")
        if table_id is None:
            return {"error": "convert_to_range 需要 table_index 或 table_name"}
        try:
            lo = sheet.ListObjects(table_id if isinstance(table_id, str) else int(table_id))
            name = lo.Name
            lo.Unlist()
            results.append(f"表格 '{name}' 已转回普通区域")
        except Exception as exc:
            return {"error": f"转换为区域失败: {exc}"}

    # ============================================================
    # 批次 D.1：数据操作
    # ============================================================

    # ===== 删除重复值 =====
    elif action == "remove_duplicates":
        target_range = params.get("range")
        if not target_range:
            return {"error": "remove_duplicates 需要 range 参数"}
        has_header = params.get("has_header", True)
        columns = params.get("columns")
        try:
            rng = sheet.Range(target_range)
            if columns is None:
                col_list = list(range(1, rng.Columns.Count + 1))
            elif isinstance(columns, (list, tuple)):
                col_list = [int(x) for x in columns]
            else:
                col_list = [int(columns)]
            header_flag = 1 if has_header else 2
            rows_before = rng.Rows.Count
            rng.RemoveDuplicates(Columns=col_list, Header=header_flag)
            results.append(f"{target_range} 已删除重复值（按列 {col_list}，{rows_before} → {rng.Rows.Count} 行）")
        except Exception as exc:
            return {"error": f"删除重复值失败: {exc}"}

    # ===== 文本分列 =====
    elif action == "text_to_columns":
        target_range = params.get("range") or params.get("source")
        if not target_range:
            return {"error": "text_to_columns 需要 range 参数"}
        destination = params.get("destination")
        data_type = params.get("data_type") or params.get("parse_type", "delimited")
        data_type_val = TEXT_PARSING_TYPE_MAP.get(str(data_type).lower(), data_type)
        try:
            rng = sheet.Range(target_range)
            kwargs = {"DataType": int(data_type_val)}
            if destination:
                kwargs["Destination"] = sheet.Range(destination)
            if int(data_type_val) == 1:
                kwargs["Tab"] = bool(params.get("tab", False))
                kwargs["Semicolon"] = bool(params.get("semicolon", False))
                kwargs["Comma"] = bool(params.get("comma", False))
                kwargs["Space"] = bool(params.get("space", False))
                if params.get("other"):
                    kwargs["Other"] = True
                    kwargs["OtherChar"] = str(params["other"])
                kwargs["ConsecutiveDelimiter"] = bool(params.get("consecutive_delimiter", False))
                tq = params.get("text_qualifier")
                if tq is not None:
                    tq_map = {"double": 1, "none": 2, "single": 3, "双引号": 1, "无": 2, "单引号": 3}
                    kwargs["TextQualifier"] = tq_map.get(str(tq).lower(), 1)
            field_info = params.get("field_info")
            if isinstance(field_info, list):
                kwargs["FieldInfo"] = [list(x) if isinstance(x, (list, tuple)) else [x] for x in field_info]
            rng.TextToColumns(**kwargs)
            results.append(f"{target_range} 已完成文本分列（{data_type}）")
        except Exception as exc:
            return {"error": f"文本分列失败: {exc}"}

    # ===== 填充序列 =====
    elif action == "fill_series":
        target_range = params.get("range") or params.get("destination")
        if not target_range:
            return {"error": "fill_series 需要 range 参数"}
        series_type = params.get("type", "linear")
        type_val = SERIES_TYPE_MAP.get(str(series_type).lower(), series_type)
        date_unit = params.get("date_unit", "day")
        date_unit_val = SERIES_DATE_UNIT_MAP.get(str(date_unit).lower(), date_unit)
        step = params.get("step", 1)
        stop = params.get("stop")
        rowcol = params.get("rowcol", "columns")
        rowcol_val = 1 if str(rowcol).lower() in ("rows", "行") else 2
        try:
            rng = sheet.Range(target_range)
            kwargs = {
                "Rowcol": rowcol_val, "Type": int(type_val), "Date": int(date_unit_val),
                "Step": float(step), "Trend": bool(params.get("trend", False)),
            }
            if stop is not None:
                kwargs["Stop"] = float(stop)
            rng.DataSeries(**kwargs)
            results.append(f"{target_range} 已按 {series_type} 填充序列（步长 {step}）")
        except Exception as exc:
            return {"error": f"填充序列失败: {exc}"}

    # ===== 自动填充 =====
    elif action == "auto_fill":
        source = params.get("source")
        destination = params.get("destination")
        if not source or not destination:
            return {"error": "auto_fill 需要 source 和 destination"}
        af_map = {
            "default": 0, "默认": 0, "copy": 1, "复制": 1, "series": 2, "序列": 2,
            "format": 3, "仅格式": 3, "values": 4, "仅值": 4,
            "days": 5, "天": 5, "weekdays": 6, "工作日": 6,
            "months": 7, "月": 7, "years": 8, "年": 8,
            "linear_trend": 9, "线性趋势": 9, "growth_trend": 10, "增长趋势": 10,
            "flash_fill": 11, "快速填充": 11,
        }
        af_type = params.get("fill_type", "default")
        af_val = af_map.get(str(af_type).lower(), af_type)
        try:
            sheet.Range(source).AutoFill(Destination=sheet.Range(destination), Type=int(af_val))
            results.append(f"自动填充 {source} → {destination}（{af_type}）")
        except Exception as exc:
            return {"error": f"自动填充失败: {exc}"}

    # ===== 分类汇总 =====
    elif action == "subtotal":
        target_range = params.get("range")
        if not target_range:
            return {"error": "subtotal 需要 range 参数"}
        group_by = params.get("group_by") or params.get("group_field")
        if group_by is None:
            return {"error": "subtotal 需要 group_by 参数（列号，从 1 开始）"}
        function = params.get("function", "sum")
        func_map = {
            "average": -4106, "平均": -4106, "avg": -4106,
            "count": -4112, "计数": -4112,
            "count_nums": -4113, "数值计数": -4113,
            "max": -4136, "最大": -4136,
            "min": -4139, "最小": -4139,
            "product": -4149, "乘积": -4149,
            "std_dev": -4097, "标准差": -4097, "std_dev_p": -4099,
            "sum": -4157, "求和": -4157, "合计": -4157,
            "var": -4164, "方差": -4164, "var_p": -4165,
        }
        func_val = func_map.get(str(function).lower(), function)
        totals_cols = params.get("totals_columns") or params.get("totals")
        if not isinstance(totals_cols, (list, tuple)):
            return {"error": "subtotal 需要 totals_columns 数组"}
        try:
            sheet.Range(target_range).Subtotal(
                GroupBy=int(group_by), Function=int(func_val),
                TotalList=[int(x) for x in totals_cols],
                Replace=bool(params.get("replace", True)),
                PageBreaks=bool(params.get("page_breaks", False)),
                SummaryBelowData=1 if params.get("summary_below", True) else 0)
            results.append(f"{target_range} 已按列 {group_by} 分类汇总（{function}）")
        except Exception as exc:
            return {"error": f"分类汇总失败: {exc}"}

    # ===== 删除分类汇总 =====
    elif action == "remove_subtotal":
        target_range = params.get("range")
        if not target_range:
            return {"error": "remove_subtotal 需要 range 参数"}
        try:
            sheet.Range(target_range).RemoveSubtotal()
            results.append(f"{target_range} 分类汇总已删除")
        except Exception as exc:
            return {"error": f"删除分类汇总失败: {exc}"}

    # ===== 分组 =====
    elif action == "group":
        try:
            if params.get("rows"):
                sheet.Rows(params["rows"]).Group()
                results.append(f"行 {params['rows']} 已分组")
            elif params.get("cols"):
                sheet.Columns(params["cols"]).Group()
                results.append(f"列 {params['cols']} 已分组")
            elif params.get("range"):
                sheet.Range(params["range"]).Group()
                results.append(f"{params['range']} 已分组")
            else:
                return {"error": "group 需要 rows / cols / range 之一"}
        except Exception as exc:
            return {"error": f"分组失败: {exc}"}

    # ===== 取消分组 =====
    elif action == "ungroup":
        try:
            if params.get("rows"):
                sheet.Rows(params["rows"]).Ungroup()
                results.append(f"行 {params['rows']} 已取消分组")
            elif params.get("cols"):
                sheet.Columns(params["cols"]).Ungroup()
                results.append(f"列 {params['cols']} 已取消分组")
            elif params.get("range"):
                sheet.Range(params["range"]).Ungroup()
                results.append(f"{params['range']} 已取消分组")
            else:
                sheet.Cells.ClearOutline()
                results.append("已清除所有分级显示")
        except Exception as exc:
            return {"error": f"取消分组失败: {exc}"}

    # ===== 显示指定级别分级 =====
    elif action == "show_outline_level":
        if params.get("rows") is not None:
            try:
                sheet.Outline.ShowLevels(RowLevels=int(params["rows"]))
                results.append(f"行级别 → {params['rows']}")
            except Exception as exc:
                errors.append(f"行分级失败: {exc}")
        if params.get("cols") is not None:
            try:
                sheet.Outline.ShowLevels(ColumnLevels=int(params["cols"]))
                results.append(f"列级别 → {params['cols']}")
            except Exception as exc:
                errors.append(f"列分级失败: {exc}")

    # ============================================================
    # 批次 D.2：完整数据验证（覆盖 list/whole/decimal/date/time/text_length/custom）
    # ============================================================
    elif action == "data_validation_v2":
        target = params.get("range") or params.get("cell")
        if not target:
            return {"error": "data_validation_v2 需要 range 或 cell 参数"}
        v_type = params.get("validation_type", "list")
        vtype_map = {
            "any": 0, "whole_number": 1, "整数": 1, "whole": 1,
            "decimal": 2, "小数": 2,
            "list": 3, "序列": 3,
            "date": 4, "日期": 4,
            "time": 5, "时间": 5,
            "text_length": 6, "文本长度": 6,
            "custom": 7, "自定义": 7,
        }
        vt = vtype_map.get(str(v_type).lower(), v_type)
        op_map = {
            "between": 1, "之间": 1,
            "not_between": 2, "未介于": 2,
            "equal": 3, "等于": 3, "=": 3,
            "not_equal": 4, "不等于": 4, "!=": 4, "<>": 4,
            "greater": 5, "大于": 5, ">": 5,
            "less": 6, "小于": 6, "<": 6,
            "greater_equal": 7, "大于等于": 7, ">=": 7,
            "less_equal": 8, "小于等于": 8, "<=": 8,
        }
        op = op_map.get(str(params.get("operator", "between")).lower(), 1)
        alert_map = {"stop": 1, "停止": 1, "warning": 2, "警告": 2, "info": 3, "信息": 3}
        alert_style = alert_map.get(str(params.get("alert_style", "stop")).lower(), 1)
        formula1 = params.get("formula1")
        if formula1 is None and isinstance(params.get("items"), (list, tuple)):
            formula1 = ",".join(str(x) for x in params["items"])
        formula2 = params.get("formula2")
        try:
            rng = sheet.Range(target)
            try:
                rng.Validation.Delete()
            except Exception:
                pass
            add_kwargs = {"Type": int(vt), "AlertStyle": int(alert_style), "Operator": int(op)}
            if formula1 is not None:
                add_kwargs["Formula1"] = str(formula1)
            if formula2 is not None:
                add_kwargs["Formula2"] = str(formula2)
            rng.Validation.Add(**add_kwargs)
            if params.get("input_title") or params.get("input_message"):
                try:
                    rng.Validation.InputTitle = str(params.get("input_title", ""))
                    rng.Validation.InputMessage = str(params.get("input_message", ""))
                    rng.Validation.ShowInput = True
                except Exception:
                    pass
            if params.get("error_title") or params.get("error_message"):
                try:
                    rng.Validation.ErrorTitle = str(params.get("error_title", ""))
                    rng.Validation.ErrorMessage = str(params.get("error_message", ""))
                    rng.Validation.ShowError = True
                except Exception:
                    pass
            try:
                rng.Validation.IgnoreBlank = bool(params.get("ignore_blank", True))
                if int(vt) == 3:
                    rng.Validation.InCellDropdown = bool(params.get("in_cell_dropdown", True))
            except Exception:
                pass
            results.append(f"{target} 数据验证已设置（{v_type}）")
        except Exception as exc:
            return {"error": f"设置数据验证失败: {exc}"}

    # ===== 圈出无效数据 =====
    elif action == "circle_invalid_data":
        try:
            wb.Application.ActiveSheet.CircleInvalid()
            results.append("已圈出无效数据")
        except Exception as exc:
            return {"error": f"圈出无效数据失败: {exc}"}

    # ===== 清除无效数据圈 =====
    elif action == "clear_invalid_circles":
        try:
            wb.Application.ActiveSheet.ClearCircles()
            results.append("无效数据圈已清除")
        except Exception as exc:
            return {"error": f"清除无效数据圈失败: {exc}"}

    # ============================================================
    # 批次 D.3：条件格式扩展（top/bottom/above_average/text_contains 等）
    # ============================================================
    elif action == "conditional_format_v2":
        target = params.get("range") or params.get("cell")
        if not target:
            return {"error": "conditional_format_v2 需要 range 参数"}
        cf_type = params.get("cf_type") or params.get("rule_type")
        if not cf_type:
            return {"error": "conditional_format_v2 需要 cf_type 参数"}
        cf_type = str(cf_type).lower()
        try:
            rng = sheet.Range(target)
            fc = None
            if cf_type in ("top", "top_n", "前", "前n"):
                fc = rng.FormatConditions.AddTop10()
                fc.TopBottom = 0
                fc.Rank = int(params.get("rank", 10))
                fc.Percent = bool(params.get("percent", False))
            elif cf_type in ("bottom", "bottom_n", "后", "后n"):
                fc = rng.FormatConditions.AddTop10()
                fc.TopBottom = 1
                fc.Rank = int(params.get("rank", 10))
                fc.Percent = bool(params.get("percent", False))
            elif cf_type in ("above_average", "高于平均", "高于平均值"):
                fc = rng.FormatConditions.AddAboveAverage()
                fc.AboveBelow = 0
                if params.get("std_dev") is not None:
                    fc.NumStdDev = int(params["std_dev"])
            elif cf_type in ("below_average", "低于平均", "低于平均值"):
                fc = rng.FormatConditions.AddAboveAverage()
                fc.AboveBelow = 1
                if params.get("std_dev") is not None:
                    fc.NumStdDev = int(params["std_dev"])
            elif cf_type in ("text_contains", "包含文本", "包含特定文本"):
                fc = rng.FormatConditions.Add(Type=9, Operator=0, String=str(params.get("text", "")))
            elif cf_type in ("text_not_contains", "不包含文本"):
                fc = rng.FormatConditions.Add(Type=9, Operator=1, String=str(params.get("text", "")))
            elif cf_type in ("text_begins_with", "开头是", "begins_with"):
                fc = rng.FormatConditions.Add(Type=9, Operator=2, String=str(params.get("text", "")))
            elif cf_type in ("text_ends_with", "结尾是", "ends_with"):
                fc = rng.FormatConditions.Add(Type=9, Operator=3, String=str(params.get("text", "")))
            elif cf_type in ("blank", "空值", "空单元格"):
                fc = rng.FormatConditions.Add(Type=10)  # xlBlanksCondition
            elif cf_type in ("no_blank", "非空", "非空白"):
                fc = rng.FormatConditions.Add(Type=13)  # xlNoBlanksCondition
            elif cf_type in ("error", "错误", "包含错误"):
                fc = rng.FormatConditions.Add(Type=16)  # xlErrorsCondition
            elif cf_type in ("no_error", "无错误"):
                fc = rng.FormatConditions.Add(Type=17)  # xlNoErrorsCondition
            elif cf_type in ("time_period", "日期发生", "日期期间"):
                period_map = {
                    "yesterday": 0, "昨天": 0,
                    "today": 1, "今天": 1,
                    "tomorrow": 2, "明天": 2,
                    "last_7_days": 3, "最近7天": 3,
                    "last_week": 4, "上周": 4,
                    "this_week": 5, "本周": 5,
                    "next_week": 6, "下周": 6,
                    "last_month": 7, "上月": 7,
                    "this_month": 8, "本月": 8,
                    "next_month": 9, "下月": 9,
                }
                period = params.get("period", "today")
                period_val = period_map.get(str(period).lower(), period)
                fc = rng.FormatConditions.Add(Type=11, DatePeriod=int(period_val))  # xlTimePeriod
            elif cf_type in ("duplicate", "重复值"):
                fc = rng.FormatConditions.AddUniqueValues()
                fc.DupeUnique = 1  # xlDuplicate
            elif cf_type in ("unique", "唯一值"):
                fc = rng.FormatConditions.AddUniqueValues()
                fc.DupeUnique = 0  # xlUnique
            else:
                return {"error": f"未知的 cf_type: {cf_type}"}
            # 应用样式（调用模块级 _apply_cf_format）
            fmt = params.get("format") or {}
            if isinstance(fmt, str):
                try:
                    fmt = json.loads(fmt) if fmt.strip() else {}
                except Exception:
                    fmt = {}
            if fc is not None and isinstance(fmt, dict):
                _apply_cf_format(fc, fmt)
            try:
                if params.get("stop_if_true") is not None:
                    fc.StopIfTrue = bool(params["stop_if_true"])
            except Exception:
                pass
            results.append(f"{target} 已添加条件格式（{cf_type}）")
        except Exception as exc:
            return {"error": f"条件格式失败: {exc}"}

    # ===== 删除指定 index 的条件格式 =====
    elif action == "delete_conditional_format":
        target = params.get("range") or params.get("cell")
        if not target:
            return {"error": "delete_conditional_format 需要 range 参数"}
        index = params.get("index")
        try:
            rng = sheet.Range(target)
            if index is None:
                rng.FormatConditions.Delete()
                results.append(f"{target} 全部条件格式已删除")
            else:
                rng.FormatConditions(int(index)).Delete()
                results.append(f"{target} 第 {index} 条条件格式已删除")
        except Exception as exc:
            return {"error": f"删除条件格式失败: {exc}"}

    # ===== 列出当前区域的条件格式 =====
    elif action == "list_conditional_formats":
        target = params.get("range") or params.get("cell")
        if not target:
            target = sheet.UsedRange.Address
        try:
            rng = sheet.Range(target)
            cfs = []
            for i in range(1, rng.FormatConditions.Count + 1):
                cf = rng.FormatConditions(i)
                info = {"index": i}
                try: info["type_code"] = int(cf.Type)
                except Exception: pass
                try: info["formula1"] = cf.Formula1 if hasattr(cf, "Formula1") else None
                except Exception: pass
                try: info["range"] = cf.AppliesTo.Address if cf.AppliesTo else None
                except Exception: pass
                cfs.append(info)
            return {"success": True, "range": target, "count": len(cfs), "conditional_formats": cfs}
        except Exception as exc:
            return {"error": f"列出条件格式失败: {exc}"}

    # ============================================================
    # 批次 E.1：页面打印
    # ============================================================

    # ===== 页面设置 =====
    elif action == "page_setup":
        ps = sheet.PageSetup
        setup = params.get("setup") or params
        try:
            if "orientation" in setup:
                ov = _resolve_orientation(setup["orientation"])
                if ov is not None:
                    ps.Orientation = ov
            if "paper_size" in setup:
                pv = _resolve_paper_size(setup["paper_size"])
                if pv is not None:
                    ps.PaperSize = pv
            for src, dst in [
                ("top_margin", "TopMargin"), ("bottom_margin", "BottomMargin"),
                ("left_margin", "LeftMargin"), ("right_margin", "RightMargin"),
                ("header_margin", "HeaderMargin"), ("footer_margin", "FooterMargin"),
            ]:
                if src in setup and setup[src] is not None:
                    # Excel 接受 point 值（1 inch = 72 points）；如果传入数值 < 5 视为 inch
                    val = float(setup[src])
                    if val < 5:
                        val *= 72
                    setattr(ps, dst, val)
            if "zoom" in setup:
                z = setup["zoom"]
                if z is False or z == 0:
                    ps.Zoom = False
                else:
                    ps.Zoom = int(z)
            if "fit_to_pages_wide" in setup:
                ps.Zoom = False
                ps.FitToPagesWide = int(setup["fit_to_pages_wide"])
            if "fit_to_pages_tall" in setup:
                ps.Zoom = False
                ps.FitToPagesTall = int(setup["fit_to_pages_tall"])
            if "center_horizontally" in setup:
                ps.CenterHorizontally = bool(setup["center_horizontally"])
            if "center_vertically" in setup:
                ps.CenterVertically = bool(setup["center_vertically"])
            if "print_gridlines" in setup:
                ps.PrintGridlines = bool(setup["print_gridlines"])
            if "print_headings" in setup:
                ps.PrintHeadings = bool(setup["print_headings"])
            if "black_and_white" in setup:
                ps.BlackAndWhite = bool(setup["black_and_white"])
            if "draft" in setup:
                ps.Draft = bool(setup["draft"])
            if "order" in setup:
                # xlDownThenOver=1, xlOverThenDown=2
                order_map = {"down_then_over": 1, "先列后行": 1, "先下后右": 1,
                             "over_then_down": 2, "先行后列": 2, "先右后下": 2}
                ov = order_map.get(str(setup["order"]).lower(), setup["order"])
                ps.Order = int(ov)
            results.append("页面设置已应用")
        except Exception as exc:
            return {"error": f"页面设置失败: {exc}"}

    # ===== 设置打印区域 =====
    elif action == "print_area":
        area = params.get("range") or params.get("area") or ""
        try:
            sheet.PageSetup.PrintArea = str(area) if area else ""
            if area:
                results.append(f"打印区域 → {area}")
            else:
                results.append("打印区域已清除")
        except Exception as exc:
            return {"error": f"设置打印区域失败: {exc}"}

    # ===== 设置打印标题 =====
    elif action == "print_titles":
        try:
            if "rows" in params:
                sheet.PageSetup.PrintTitleRows = str(params["rows"]) if params["rows"] else ""
                results.append(f"标题行 → {params['rows']}")
            if "cols" in params or "columns" in params:
                cols = params.get("cols") or params.get("columns")
                sheet.PageSetup.PrintTitleColumns = str(cols) if cols else ""
                results.append(f"标题列 → {cols}")
        except Exception as exc:
            return {"error": f"设置打印标题失败: {exc}"}

    # ===== 分页符 =====
    elif action == "page_break":
        op = params.get("operation", "add")  # add / remove / remove_all
        cell = params.get("cell")
        try:
            if op == "remove_all":
                sheet.ResetAllPageBreaks()
                results.append("所有分页符已清除")
            elif op == "remove":
                if not cell:
                    return {"error": "page_break remove 需要 cell"}
                c = sheet.Range(cell)
                if params.get("type") in ("horizontal", "水平"):
                    sheet.HPageBreaks(int(params.get("index", 1))).Delete()
                elif params.get("type") in ("vertical", "垂直"):
                    sheet.VPageBreaks(int(params.get("index", 1))).Delete()
                else:
                    # 同时清除该 cell 的水平和垂直分页符
                    try: c.PageBreak = -4142  # xlPageBreakNone
                    except Exception: pass
                results.append(f"分页符已删除 ({cell})")
            else:
                # add
                if not cell:
                    return {"error": "page_break add 需要 cell"}
                c = sheet.Range(cell)
                # xlPageBreakManual=-4135
                c.PageBreak = -4135
                results.append(f"已在 {cell} 添加分页符")
        except Exception as exc:
            return {"error": f"分页符操作失败: {exc}"}

    # ===== 页眉页脚 =====
    elif action == "set_header" or action == "set_footer":
        ps = sheet.PageSetup
        is_header = (action == "set_header")
        try:
            if "left" in params:
                if is_header:
                    ps.LeftHeader = str(params["left"])
                else:
                    ps.LeftFooter = str(params["left"])
            if "center" in params:
                if is_header:
                    ps.CenterHeader = str(params["center"])
                else:
                    ps.CenterFooter = str(params["center"])
            if "right" in params:
                if is_header:
                    ps.RightHeader = str(params["right"])
                else:
                    ps.RightFooter = str(params["right"])
            # 单参数风格
            if "text" in params:
                if is_header:
                    ps.CenterHeader = str(params["text"])
                else:
                    ps.CenterFooter = str(params["text"])
            results.append(f"{'页眉' if is_header else '页脚'}已设置")
        except Exception as exc:
            return {"error": f"设置页眉页脚失败: {exc}"}

    # ===== 打印预览 / 实际打印 =====
    elif action == "print_preview":
        try:
            sheet.PrintPreview()
            results.append("已打开打印预览")
        except Exception as exc:
            return {"error": f"打印预览失败: {exc}"}

    # ============================================================
    # 批次 E.2：工作表外观
    # ============================================================

    # ===== 工作表标签颜色 =====
    elif action == "set_tab_color":
        color = params.get("color")
        try:
            if color is None or color == "":
                sheet.Tab.ColorIndex = -4142  # xlColorIndexNone
                results.append("标签颜色已清除")
            else:
                sheet.Tab.Color = color_to_ole(color)
                results.append(f"标签颜色 → {color}")
        except Exception as exc:
            return {"error": f"设置标签颜色失败: {exc}"}

    # ===== 缩放级别 =====
    elif action == "set_zoom":
        zoom = params.get("zoom", 100)
        try:
            # 需要激活工作表
            sheet.Activate()
            excel.ActiveWindow.Zoom = int(zoom)
            results.append(f"缩放 → {zoom}%")
        except Exception as exc:
            return {"error": f"设置缩放失败: {exc}"}

    # ===== 显示/隐藏网格线 =====
    elif action == "set_gridlines_visible":
        visible = params.get("visible", True)
        try:
            sheet.Activate()
            excel.ActiveWindow.DisplayGridlines = bool(visible)
            results.append(f"网格线 → {'显示' if visible else '隐藏'}")
        except Exception as exc:
            return {"error": f"切换网格线失败: {exc}"}

    # ===== 显示/隐藏行列标题 =====
    elif action == "set_headings_visible":
        visible = params.get("visible", True)
        try:
            sheet.Activate()
            excel.ActiveWindow.DisplayHeadings = bool(visible)
            results.append(f"行列标题 → {'显示' if visible else '隐藏'}")
        except Exception as exc:
            return {"error": f"切换行列标题失败: {exc}"}

    # ===== 移动工作表顺序 =====
    elif action == "move_sheet":
        try:
            before_name = params.get("before")
            after_name = params.get("after")
            position = params.get("position")  # 数字索引
            if before_name:
                sheet.Move(Before=wb.Sheets(before_name))
                results.append(f"工作表已移至 '{before_name}' 之前")
            elif after_name:
                sheet.Move(After=wb.Sheets(after_name))
                results.append(f"工作表已移至 '{after_name}' 之后")
            elif position is not None:
                pos = int(position)
                if pos <= 1:
                    sheet.Move(Before=wb.Sheets(1))
                elif pos >= wb.Sheets.Count:
                    sheet.Move(After=wb.Sheets(wb.Sheets.Count))
                else:
                    sheet.Move(Before=wb.Sheets(pos))
                results.append(f"工作表已移至位置 {pos}")
            else:
                return {"error": "move_sheet 需要 before/after/position 之一"}
        except Exception as exc:
            return {"error": f"移动工作表失败: {exc}"}

    # ===== 拆分窗口 =====
    elif action == "split":
        cell = params.get("cell") or params.get("split_at")
        try:
            sheet.Activate()
            if cell:
                excel.ActiveWindow.SplitColumn = sheet.Range(cell).Column - 1
                excel.ActiveWindow.SplitRow = sheet.Range(cell).Row - 1
            elif "split_row" in params or "split_column" in params:
                if "split_row" in params:
                    excel.ActiveWindow.SplitRow = int(params["split_row"])
                if "split_column" in params:
                    excel.ActiveWindow.SplitColumn = int(params["split_column"])
            excel.ActiveWindow.Split = True
            results.append("窗口已拆分")
        except Exception as exc:
            return {"error": f"拆分窗口失败: {exc}"}

    # ===== 取消拆分 =====
    elif action == "unsplit":
        try:
            sheet.Activate()
            excel.ActiveWindow.Split = False
            results.append("已取消拆分")
        except Exception as exc:
            return {"error": f"取消拆分失败: {exc}"}

    # ===== 计算控制 =====
    elif action == "calculate":
        scope = params.get("scope", "sheet")  # workbook / sheet / range
        try:
            if scope == "workbook":
                wb.Application.Calculate()
                results.append("整个工作簿已重算")
            elif scope == "range" and params.get("range"):
                sheet.Range(params["range"]).Calculate()
                results.append(f"{params['range']} 已重算")
            else:
                sheet.Calculate()
                results.append(f"工作表 '{sheet.Name}' 已重算")
        except Exception as exc:
            return {"error": f"重算失败: {exc}"}

    # ===== 设置计算模式 =====
    elif action == "set_calculation_mode":
        mode = params.get("mode", "automatic")
        mode_map = {
            "automatic": -4105, "自动": -4105,
            "manual": -4135, "手动": -4135,
            "semi_automatic": 2, "半自动": 2, "自动重算除模拟运算外": 2,
        }
        mode_val = mode_map.get(str(mode).lower(), mode)
        try:
            excel.Calculation = int(mode_val)
            results.append(f"计算模式 → {mode}")
        except Exception as exc:
            return {"error": f"设置计算模式失败: {exc}"}

    # ============================================================
    # 批次 F：形状/文本框/图片管理
    # ============================================================

    # ===== 列出工作表所有 Shape（不含图表）=====
    elif action == "list_shapes":
        items = _read_excel_shapes(sheet, max_shapes=int(params.get("limit", 200)))
        return {"success": True, "sheet": sheet.Name, "count": len(items), "shapes": items}

    # ===== 插入形状 =====
    elif action == "insert_shape":
        shape_type = params.get("shape_type") or params.get("type", "rectangle")
        shape_val = AUTOSHAPE_TYPE_MAP.get(str(shape_type).lower(), shape_type)
        try:
            shape_val = int(shape_val)
        except Exception:
            return {"error": f"无法解析形状类型: {shape_type}"}
        # 位置/大小：优先支持 cell+宽高，其次 left/top
        left = params.get("left")
        top = params.get("top")
        width = float(params.get("width", 100))
        height = float(params.get("height", 50))
        if params.get("cell"):
            anchor = sheet.Range(params["cell"])
            left = float(anchor.Left)
            top = float(anchor.Top)
        if left is None or top is None:
            return {"error": "insert_shape 需要 cell 或 left+top 定位"}
        try:
            if shape_val == 9999:
                # 直线
                shp = sheet.Shapes.AddLine(float(left), float(top),
                                           float(left) + width, float(top) + height)
            else:
                shp = sheet.Shapes.AddShape(shape_val, float(left), float(top), width, height)
            # 名称
            if params.get("name"):
                shp.Name = str(params["name"])
            # 文本
            if params.get("text") and hasattr(shp, "TextFrame2"):
                try:
                    shp.TextFrame2.TextRange.Text = str(params["text"])
                except Exception:
                    try: shp.TextFrame.Characters().Text = str(params["text"])
                    except Exception: pass
            # 填充
            if params.get("fill_color") is not None:
                try:
                    shp.Fill.Visible = -1
                    shp.Fill.Solid()
                    shp.Fill.ForeColor.RGB = color_to_ole(params["fill_color"])
                except Exception:
                    pass
            if params.get("border_color") is not None:
                try:
                    shp.Line.Visible = -1
                    shp.Line.ForeColor.RGB = color_to_ole(params["border_color"])
                except Exception:
                    pass
            if params.get("border_width") is not None:
                try:
                    shp.Line.Weight = float(params["border_width"])
                except Exception:
                    pass
            if isinstance(params.get("font"), dict) and hasattr(shp, "TextFrame2"):
                try:
                    _apply_text_font(shp.TextFrame2.TextRange.Font, params["font"])
                except Exception:
                    pass
            if params.get("rotation") is not None:
                try: shp.Rotation = float(params["rotation"])
                except Exception: pass
            results.append(f"已插入形状 '{shp.Name}'（{shape_type}）")
        except Exception as exc:
            return {"error": f"插入形状失败: {exc}"}

    # ===== 插入文本框 =====
    elif action == "insert_textbox":
        text = params.get("text", "")
        left = params.get("left")
        top = params.get("top")
        width = float(params.get("width", 200))
        height = float(params.get("height", 80))
        if params.get("cell"):
            anchor = sheet.Range(params["cell"])
            left = float(anchor.Left)
            top = float(anchor.Top)
        if left is None or top is None:
            return {"error": "insert_textbox 需要 cell 或 left+top"}
        try:
            # msoTextOrientationHorizontal = 1
            shp = sheet.Shapes.AddTextbox(1, float(left), float(top), width, height)
            if params.get("name"):
                shp.Name = str(params["name"])
            shp.TextFrame2.TextRange.Text = str(text)
            if isinstance(params.get("font"), dict):
                _apply_text_font(shp.TextFrame2.TextRange.Font, params["font"])
            if params.get("fill_color") is not None:
                try:
                    shp.Fill.Visible = -1
                    shp.Fill.Solid()
                    shp.Fill.ForeColor.RGB = color_to_ole(params["fill_color"])
                except Exception:
                    pass
            if params.get("border_color") is not None:
                try:
                    shp.Line.Visible = -1
                    shp.Line.ForeColor.RGB = color_to_ole(params["border_color"])
                except Exception:
                    pass
            if params.get("border_visible") is False:
                try: shp.Line.Visible = 0
                except Exception: pass
            # 文本对齐
            align_map = {"left": 1, "center": 2, "right": 3, "justify": 4,
                         "左": 1, "居中": 2, "中": 2, "右": 3, "两端对齐": 4}
            if params.get("text_align"):
                try:
                    av = align_map.get(str(params["text_align"]).lower(), params["text_align"])
                    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = int(av)
                except Exception:
                    pass
            results.append(f"已插入文本框 '{shp.Name}'")
        except Exception as exc:
            return {"error": f"插入文本框失败: {exc}"}

    # ===== 删除形状 =====
    elif action == "delete_shape":
        identifier = params.get("name") or params.get("index")
        # 兼容：传 chart_index 时通过图表查找定位 Shape
        if identifier is None and params.get("chart_index") is not None:
            ch = _find_chart_by_identifier(sheet, params["chart_index"])
            if ch is not None:
                identifier = ch.Name
        if identifier is None:
            return {"error": "delete_shape 需要 name/index/chart_index 之一"}
        try:
            if isinstance(identifier, str):
                sheet.Shapes(identifier).Delete()
            else:
                sheet.Shapes(int(identifier)).Delete()
            results.append(f"形状 '{identifier}' 已删除")
        except Exception as exc:
            return {"error": f"删除形状失败: {exc}"}

    # ===== 编辑/移动形状 =====
    elif action == "edit_shape":
        identifier = params.get("name") or params.get("index")
        if identifier is None and params.get("chart_index") is not None:
            ch = _find_chart_by_identifier(sheet, params["chart_index"])
            if ch is not None:
                identifier = ch.Name
        if identifier is None:
            return {"error": "edit_shape 需要 name/index/chart_index 之一"}
        try:
            if isinstance(identifier, str):
                shp = sheet.Shapes(identifier)
            else:
                shp = sheet.Shapes(int(identifier))
        except Exception as exc:
            return {"error": f"未找到形状: {identifier}（{exc}）"}
        if params.get("left") is not None: shp.Left = float(params["left"])
        if params.get("top") is not None: shp.Top = float(params["top"])
        if params.get("width") is not None: shp.Width = float(params["width"])
        if params.get("height") is not None: shp.Height = float(params["height"])
        if params.get("rotation") is not None:
            try: shp.Rotation = float(params["rotation"])
            except Exception: pass
        if "text" in params and hasattr(shp, "TextFrame2"):
            try: shp.TextFrame2.TextRange.Text = str(params["text"])
            except Exception: pass
        if params.get("fill_color") is not None:
            try:
                shp.Fill.Visible = -1
                shp.Fill.Solid()
                shp.Fill.ForeColor.RGB = color_to_ole(params["fill_color"])
            except Exception: pass
        if params.get("border_color") is not None:
            try:
                shp.Line.Visible = -1
                shp.Line.ForeColor.RGB = color_to_ole(params["border_color"])
            except Exception: pass
        if params.get("border_width") is not None:
            try: shp.Line.Weight = float(params["border_width"])
            except Exception: pass
        if isinstance(params.get("font"), dict) and hasattr(shp, "TextFrame2"):
            try: _apply_text_font(shp.TextFrame2.TextRange.Font, params["font"])
            except Exception: pass
        results.append(f"形状 '{shp.Name}' 已更新")

    # ===== 替换图片（保留位置/大小）=====
    elif action == "replace_image":
        identifier = params.get("name") or params.get("index")
        if identifier is None and params.get("chart_index") is not None:
            ch = _find_chart_by_identifier(sheet, params["chart_index"])
            if ch is not None:
                identifier = ch.Name
        new_path = params.get("path") or params.get("new_path")
        if identifier is None or not new_path:
            return {"error": "replace_image 需要 name/index/chart_index 和 path"}
        try:
            old_shape = sheet.Shapes(identifier) if isinstance(identifier, str) else sheet.Shapes(int(identifier))
        except Exception as exc:
            return {"error": f"未找到图片: {identifier}（{exc}）"}
        left = float(old_shape.Left)
        top = float(old_shape.Top)
        width = float(old_shape.Width)
        height = float(old_shape.Height)
        old_name = old_shape.Name
        old_shape.Delete()
        try:
            # 下载/解析路径（兼容本地路径和远程 URL）
            image_path, _temp_path, image_error = _prepare_office_image_path(new_path)
            if image_error:
                return {"error": image_error}
            new_shape = sheet.Shapes.AddPicture(Filename=image_path,
                                                LinkToFile=False, SaveWithDocument=True,
                                                Left=left, Top=top, Width=width, Height=height)
            new_shape.Name = old_name
            results.append(f"图片 '{old_name}' 已替换为 {new_path}")
        except Exception as exc:
            return {"error": f"替换图片失败: {exc}"}

    # ===== 移动图片（专用，复用 edit_shape 实现）=====
    elif action == "move_shape":
        identifier = params.get("name") or params.get("index")
        if identifier is None and params.get("chart_index") is not None:
            ch = _find_chart_by_identifier(sheet, params["chart_index"])
            if ch is not None:
                identifier = ch.Name
        if identifier is None:
            return {"error": "move_shape 需要 name/index/chart_index 之一"}
        try:
            shp = sheet.Shapes(identifier) if isinstance(identifier, str) else sheet.Shapes(int(identifier))
            if params.get("cell"):
                anchor = sheet.Range(params["cell"])
                shp.Left = float(anchor.Left)
                shp.Top = float(anchor.Top)
            else:
                if params.get("left") is not None: shp.Left = float(params["left"])
                if params.get("top") is not None: shp.Top = float(params["top"])
            results.append(f"形状 '{shp.Name}' 已移动")
        except Exception as exc:
            return {"error": f"移动形状失败: {exc}"}

    # ============================================================
    # 工作簿级 / 杂项
    # ============================================================

    # ===== 设置工作簿属性 =====
    elif action == "set_workbook_property":
        prop_map = {
            "title": "Title", "标题": "Title",
            "author": "Author", "作者": "Author",
            "subject": "Subject", "主题": "Subject",
            "keywords": "Keywords", "关键字": "Keywords",
            "comments": "Comments", "备注": "Comments",
            "category": "Category", "类别": "Category",
            "manager": "Manager", "经理": "Manager",
            "company": "Company", "公司": "Company",
        }
        for k, v in params.items():
            if k in ("filename", "edit_action", "sheet", "sheet_name", "tool"):
                continue
            if k in prop_map:
                try:
                    wb.BuiltinDocumentProperties(prop_map[k]).Value = str(v)
                    results.append(f"属性 {k} → {v}")
                except Exception as exc:
                    errors.append(f"设置属性 {k} 失败: {exc}")

    # ===== 公式追踪：追踪引用单元格 =====
    elif action == "trace_precedents":
        cell = params.get("cell")
        if not cell:
            return {"error": "trace_precedents 需要 cell"}
        try:
            sheet.Range(cell).ShowPrecedents()
            results.append(f"已显示 {cell} 的引用单元格箭头")
        except Exception as exc:
            return {"error": f"追踪引用失败: {exc}"}

    # ===== 公式追踪：追踪从属单元格 =====
    elif action == "trace_dependents":
        cell = params.get("cell")
        if not cell:
            return {"error": "trace_dependents 需要 cell"}
        try:
            sheet.Range(cell).ShowDependents()
            results.append(f"已显示 {cell} 的从属单元格箭头")
        except Exception as exc:
            return {"error": f"追踪从属失败: {exc}"}

    # ===== 清除追踪箭头 =====
    elif action == "clear_arrows":
        try:
            sheet.ClearArrows()
            results.append("追踪箭头已清除")
        except Exception as exc:
            return {"error": f"清除追踪箭头失败: {exc}"}

    else:
        return {"error": f"未知的 edit_action: {action}"}

    response = {"success": len(errors) == 0, "results": results}
    if errors:
        response["errors"] = errors
    return response


def excel_manage_sheets(params):
    """工作表管理"""
    excel = win32.GetActiveObject("Excel.Application")
    wb = find_excel_workbook(excel, params["filename"])
    if not wb:
        return {"error": f"未找到工作簿: {params['filename']}"}

    action = params.get("sheet_action", "list")
    results = []

    if action == "add":
        name = params.get("name", "")
        new_sheet = wb.Sheets.Add()
        if name:
            new_sheet.Name = name
        results.append(f"新工作表 '{new_sheet.Name}' 已创建")

    elif action == "delete":
        name = params["name"]
        excel.DisplayAlerts = False
        wb.Sheets(name).Delete()
        excel.DisplayAlerts = True
        results.append(f"工作表 '{name}' 已删除")

    elif action == "rename":
        old_name = params["old_name"]
        new_name = params["new_name"]
        wb.Sheets(old_name).Name = new_name
        results.append(f"工作表 '{old_name}' → '{new_name}'")

    elif action == "copy":
        name = params["name"]
        after = params.get("after")
        if after:
            wb.Sheets(name).Copy(After=wb.Sheets(after))
        else:
            wb.Sheets(name).Copy(After=wb.Sheets(wb.Sheets.Count))
        results.append(f"工作表 '{name}' 已复制")

    elif action == "hide":
        name = params["name"]
        wb.Sheets(name).Visible = False
        results.append(f"工作表 '{name}' 已隐藏")

    elif action == "show":
        name = params["name"]
        wb.Sheets(name).Visible = True
        results.append(f"工作表 '{name}' 已显示")

    elif action == "activate":
        name = params["name"]
        wb.Sheets(name).Activate()
        results.append(f"已切换到工作表 '{name}'")

    else:
        return {"error": f"未知的 sheet_action: {action}"}

    return {"success": True, "results": results}