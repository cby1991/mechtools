"""集中配置：文件类型、Excel 列位、布局常量。

为什么单独一个模块
------------------
原来两个脚本各自在文件头散着一堆魔数（``MODEL_EXTS``、``NAME_COL`` …），
改一处忘一处。这里集中定义，测试与工具共用同一份真相。

**本项目所有常量集中在此，业务模块不得再写裸数字。**
"""

from __future__ import annotations

from pathlib import Path
from typing import Final

__all__ = [
    "PROJECT_ROOT",
    "MODEL_EXTS",
    "DRAW_EXTS",
    "ALL_SCAN_EXTS",
    "SKIP_PREFIXES",
    "HEADER_PREFIXES",
    # --- 采购清单工作簿 ---
    "SHEET_TITLE",
    "LEGACY_SHEET_TITLE",
    "SUPPLEMENT_HEADERS",
    "SUPPLEMENT_WIDTHS",
    "COL_NO",
    "COL_IMAGE",
    "COL_NAME",
    "COL_MATERIAL",
    "COL_QTY",
    "COL_COLOR",
    "COL_CRAFT",
    "COL_NOTE",
    "COL_UNIT_PRICE",
    "COL_TOTAL_PRICE",
    "HEADER_ROW",
    "DATA_START_ROW",
    "MAX_ROWS",
    "MIN_DATA_ROWS",
    "DEFAULT_QTY",
    "HEADER_ROW_HEIGHT",
    "DATA_ROW_HEIGHT",
    "PIXELS_PER_WIDTH_UNIT",
    "PIXELS_PER_POINT",
    "CENTERED_COLUMNS",
    "HEADER_FILL",
    "HEADER_FONT_COLOR",
    "BODY_FONT_NAME",
    "HEADER_FONT_NAME",
    # --- 清单读取 ---
    "NAME_HEADER_KEYWORD",
    "HEADER_SEARCH_ROWS",
    "PREFERRED_INPUT_SHEETS",
    # --- 匹配 ---
    "DEFAULT_MIN_SCORE",
    "LOOSE_MIN_SCORE",
    "STRICT_MIN_SCORE",
    # --- 材料识别 ---
    "HEADER_ZONE_RATIO",
    # --- 技术要求提取 ---
    "TECH_TITLE_KEYWORDS",
    "TECH_ALIGN_TOLERANCE",
    "CRAFT_WORDS",
    "COLOR_WORDS",
    "CMF_MARKER",
    # --- 控制台 ---
    "RULE_WIDTH",
]

# ------------------------------------------------------------------ 路径

PROJECT_ROOT: Final[Path] = Path(__file__).resolve().parents[2]

# ------------------------------------------------------------------ 文件类型

#: 3D 模型扩展名（统一小写，比较时对扩展名做 lower）
MODEL_EXTS: Final[tuple[str, ...]] = (".stp", ".step")

#: 图纸扩展名
DRAW_EXTS: Final[tuple[str, ...]] = (".pdf",)

#: 扫描时关心的全部扩展名
ALL_SCAN_EXTS: Final[tuple[str, ...]] = MODEL_EXTS + DRAW_EXTS

#: 需要跳过的文件名前缀：Excel 临时文件 ``~$``、隐藏文件 ``.``
SKIP_PREFIXES: Final[tuple[str, ...]] = ("~$", ".")

# ------------------------------------------------------------------ 采购清单工作簿

SHEET_TITLE: Final[str] = "采购清单"

SUPPLEMENT_HEADERS: Final[tuple[str, ...]] = (
    "序号",
    "图片",
    "名称",
    "材料",
    "数量",
    "颜色",
    "成型工艺",
    "备注",
    "单价",
    "总价",
)

SUPPLEMENT_WIDTHS: Final[tuple[int, ...]] = (6, 18, 26, 12, 7, 22, 11, 30, 10, 12)

#: 列号从 1 开始（openpyxl 语义）
COL_NO: Final[int] = 1
COL_IMAGE: Final[int] = 2
COL_NAME: Final[int] = 3
COL_MATERIAL: Final[int] = 4
COL_QTY: Final[int] = 5
COL_COLOR: Final[int] = 6
COL_CRAFT: Final[int] = 7
COL_NOTE: Final[int] = 8
COL_UNIT_PRICE: Final[int] = 9
COL_TOTAL_PRICE: Final[int] = 10

HEADER_ROW: Final[int] = 1
DATA_START_ROW: Final[int] = 2
MAX_ROWS: Final[int] = 200

#: 即使没有零件，也把表格撑到这么多行，保证边框完整好看
MIN_DATA_ROWS: Final[int] = 10

DEFAULT_QTY: Final[int] = 1
HEADER_ROW_HEIGHT: Final[int] = 28

#: 数据行高。100pt 是为了给「图片」列的轴测图留出空间 ——
#: 60pt 时图片只有约 80px 高，零件细节基本糊成一团（2026-09-29 实测）。
DATA_ROW_HEIGHT: Final[int] = 100

# --- 嵌图用的像素换算（Excel 的列宽/行高不是像素，得自己换）---

#: 列宽单位 → 像素。Excel 里「列宽 1」约等于 7px（默认字体下）。
PIXELS_PER_WIDTH_UNIT: Final[int] = 7

#: 点 → 像素（96 DPI）。
PIXELS_PER_POINT: Final[float] = 96 / 72

#: 需要**水平居中**的数据列（其余列左对齐）。
#:
#: 序号、数量、单价、总价都是短数字，左对齐会贴着格线、看着散；居中才像一张表。
#: 名称、材料、备注、图片、颜色、成型工艺是长短不一的文字，居中反而难扫读，
#: 所以保持左对齐。
CENTERED_COLUMNS: Final[tuple[int, ...]] = (
    COL_NO,
    COL_QTY,
    COL_UNIT_PRICE,
    COL_TOTAL_PRICE,
)

HEADER_FILL: Final[str] = "4472C4"  # 蓝底
HEADER_FONT_COLOR: Final[str] = "FFFFFF"  # 白字
BODY_FONT_NAME: Final[str] = "宋体"
HEADER_FONT_NAME: Final[str] = "宋体"

# ------------------------------------------------------------------ 清单读取

#: 在表头里找这一列
NAME_HEADER_KEYWORD: Final[str] = "名称"

#: 表头只在前 N 行里找
HEADER_SEARCH_ROWS: Final[int] = 10

#: 改名前的旧工作表名，只用于**读**（往表里写仍然只写新名）。
LEGACY_SHEET_TITLE: Final[str] = "补料清单"

#: 读清单时**优先**认这些工作表名，按顺序匹配（靠名字，不靠位置 ——
#: 位置会因为别人插一张表就变）。
#:
#: **为什么是两个名字**：这张表 2026-09-29 之前叫「补料清单」，之后改叫
#: 「采购清单」。用户手上已有的旧文件还印着旧名 —— 只认新名的话，
#: `gbcopy` 读旧清单会退化成「取第一张表」，多表工作簿
#: （采购清单 + 备料参考）就会取错表，白拷一堆文件出去（见 AGENTS.md 的 P12）。
#: 所以**两个都认**：新名排前面，旧名兜底。
#:
#: 为什么需要它：一份工作簿里常常不止一张表带「名称」列（比如后面还跟着
#: 「备料参考」这种参考页）。早期实现把这些表的名**合并**，导致参考页
#: 的零件也被当成待采购，多拷了一堆文件出去。见 `AGENTS.md` 的 P12。
PREFERRED_INPUT_SHEETS: Final[tuple[str, ...]] = (SHEET_TITLE, LEGACY_SHEET_TITLE)

# ------------------------------------------------------------------ 匹配强度

#: 默认：完全相同 或 名称+版本号后缀（``驱动连杆`` → ``驱动连杆V1.1.step``）
DEFAULT_MIN_SCORE: Final[int] = 3

#: 宽松：互相包含也算，结果会标 ``!`` 提示人工核对
LOOSE_MIN_SCORE: Final[int] = 1

#: 严格：只认文件名与名称完全一致
STRICT_MIN_SCORE: Final[int] = 4

# ------------------------------------------------------------------ 材料识别

#: 标题栏判定：页面下方这么多比例之内的词算标题栏区域
HEADER_ZONE_RATIO: Final[float] = 0.65

# ------------------------------------------------------------------ 技术要求提取

#: 技术要求栏的标题写法。
#:
#: **为什么是一组而不是一个**：实测同一批图纸里「技术要求：」和「设计要求：」
#: 都在用（46 份里 6 份写的是「设计要求」），只认一种会整批漏掉。
TECH_TITLE_KEYWORDS: Final[tuple[str, ...]] = (
    "技术要求",
    "设计要求",
    "技术条件",
    "加工要求",
)

#: 条目与标题的**横向**对齐容差（pt）。
#:
#: 技术要求条目和标题的左边界实测完全相等（可用来把条目跟图框标题栏文字切开：
#: 后者贴在页面左边缘）。留一点余量防 PDF 导出的微小抖动，但**不能大** ——
#: 实测有一份图纸混进过一个 x0 差 7pt 的 ``0.00`` 尺寸标注。
TECH_ALIGN_TOLERANCE: Final[float] = 6.0

#: 工艺词表。表面处理 + 热处理都收 —— 对加工厂来说都是必做工序。
#:
#: **长的排在前面**：匹配按「最长优先」，否则 ``硬质阳极氧化`` 会先被
#: ``阳极氧化`` 吃掉，「硬质」这个关键信息就没了（硬质阳极氧化的膜更厚，
#: 是完全不同的报价）。
CRAFT_WORDS: Final[tuple[str, ...]] = (
    # 阳极氧化族
    "硬质阳极氧化",
    "黑色阳极氧化",
    "本色阳极氧化",
    "硫酸阳极氧化",
    "铬酸阳极氧化",
    "阳极氧化",
    "化学氧化",
    "导电氧化",
    "本色氧化",
    "铬酸盐处理",
    "无铬钝化",
    # 电镀族
    "电镀",
    "镀锌",
    "镀铬",
    "镀镍",
    "镀锡",
    "镀银",
    "镀铜",
    "镀钛",
    "镀镉",
    # 机械表面
    "喷砂",
    "喷丸",
    "抛丸",
    "喷玻璃砂",
    "镜面抛光",
    "拉丝",
    "抛光",
    "磨砂",
    "滚光",
    # 化学转化
    "发黑",
    "发蓝",
    "钝化",
    "磷化",
    "硅烷",
    # 涂装
    "喷塑",
    "喷粉",
    "喷涂",
    "喷漆",
    "电泳",
    "达克罗",
    "特氟龙",
    # 印刷 / 标记
    "丝印",
    "移印",
    "激光打标",
    "激光雕刻",
    "镭雕",
    "蚀刻",
    # 热处理
    "调质",
    "淬火",
    "回火",
    "正火",
    "退火",
    "渗碳",
    "渗氮",
    "氮化",
    "时效",
    "固溶",
    "深冷",
    "高频淬火",
    "表面淬火",
)

#: 颜色 / 外观词。
#:
#: 与工艺词分开维护，是因为**工艺词里可能已经含颜色**（``本色氧化``），
#: 这时不能再拼一个 ``本色`` 进去，否则会得到「本色氧化本色」。
COLOR_WORDS: Final[tuple[str, ...]] = (
    "本色",
    "黑色",
    "白色",
    "军绿色",
    "红色",
    "蓝色",
    "银色",
    "金色",
    "磨砂",
    "哑光",
    "亮面",
)

#: 图纸只写「表面处理参考 CMF」时的标记词（比较时不区分大小写）。
#:
#: 实测约 1/3 的图纸没有写明具体工艺，只指向 CMF 文件 —— 这类不能当「没找到」
#: 丢掉，要原样带给采购，让他知道还得去查 CMF。
CMF_MARKER: Final[str] = "CMF"

# ------------------------------------------------------------------ 控制台

RULE_WIDTH: Final[int] = 78
