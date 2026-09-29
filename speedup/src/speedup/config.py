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
    # --- 补料清单工作簿 ---
    "SHEET_TITLE",
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
    "HEADER_FILL",
    "HEADER_FONT_COLOR",
    "BODY_FONT_NAME",
    "HEADER_FONT_NAME",
    # --- 清单读取 ---
    "NAME_HEADER_KEYWORD",
    "HEADER_SEARCH_ROWS",
    "PREFERRED_INPUT_SHEET",
    # --- 匹配 ---
    "DEFAULT_MIN_SCORE",
    "LOOSE_MIN_SCORE",
    "STRICT_MIN_SCORE",
    # --- 材料识别 ---
    "HEADER_ZONE_RATIO",
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

# ------------------------------------------------------------------ 补料清单工作簿

SHEET_TITLE: Final[str] = "补料清单"

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
DATA_ROW_HEIGHT: Final[int] = 60

HEADER_FILL: Final[str] = "4472C4"  # 蓝底
HEADER_FONT_COLOR: Final[str] = "FFFFFF"  # 白字
BODY_FONT_NAME: Final[str] = "宋体"
HEADER_FONT_NAME: Final[str] = "宋体"

# ------------------------------------------------------------------ 清单读取

#: 在表头里找这一列
NAME_HEADER_KEYWORD: Final[str] = "名称"

#: 表头只在前 N 行里找
HEADER_SEARCH_ROWS: Final[int] = 10

#: 读清单时**优先**认这一张工作表，靠名字（不是靠位置）—— 位置会因为别人插一张表就变。
#: 刻意与 SHEET_TITLE 同名：自己生成的表，自己一定读得回来。
#:
#: 为什么需要它：一份工作簿里常常不止一张表带「名称」列（比如后面还跟着
#: 「备料参考」这种参考页）。早期实现把这些表的名**合并**，导致参考页
#: 的零件也被当成待补料，多拷了一堆文件出去。见 `AGENTS.md` 的 P12。
PREFERRED_INPUT_SHEET: Final[str] = SHEET_TITLE

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

# ------------------------------------------------------------------ 控制台

RULE_WIDTH: Final[int] = 78
