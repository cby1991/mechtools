"""控制台输出工具：中文宽度对齐的表格、分隔线、状态提示。

两个设计要点
------------
1. **不把 stdout 强行改成 UTF-8。** Windows 中文控制台默认 cp936(GBK)，
   强行把 Python 侧编码改成 UTF-8 会让中文直接变乱码 —— 这正是本项目历史上
   踩过的坑（见 AGENTS.md 的 R2 与 P1）。这里只把错误策略设成 ``replace``，
   保证永不抛 ``UnicodeEncodeError``。
2. **表格按显示宽度对齐。** 中日韩文字占 2 列、ASCII 占 1 列，用
   :func:`unicodedata.east_asian_width` 判定，而不是简单地区分「是否 ASCII」。
   :func:`render_table` 返回字符串而不是直接 print，方便单测断言。
"""

from __future__ import annotations

import os
import sys
import unicodedata
from collections.abc import Sequence
from typing import Any

from .config import RULE_WIDTH

__all__ = [
    "ensure_streams",
    "ensure_safe_output",
    "console_encoding",
    "display_width",
    "truncate",
    "pad",
    "render_table",
    "print_table",
    "rule",
    "heading",
    "section",
    "ok",
    "warn",
    "fail",
    "info",
]

#: east_asian_width 里这两种是「宽」字符，占 2 列
_WIDE_CATEGORIES = frozenset({"W", "F"})

ELLIPSIS = "…"


def ensure_streams() -> None:
    """补上 ``sys.stdout`` / ``sys.stderr``。

    用 ``pythonw.exe`` 启动（双击快捷方式，图省掉那个黑框控制台）时，这两个
    流是 ``None`` —— 任何一句 ``print()`` 都会抛
    ``AttributeError: 'NoneType' object has no attribute 'write'``。
    这里给它们接一个丢弃型的空流：GUI 模式下本来就不需要控制台输出，
    但诊断信息（日志框里显示的）不能因此中断。

    **必须在任何 print 之前调用。**
    """
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")  # noqa: SIM115 - 活到进程结束
    if sys.stderr is None:
        sys.stderr = sys.stdout


def ensure_safe_output() -> None:
    """让 stdout/stderr 编码失败时退化为 ``?`` 而**不是**抛异常或换编码。

    保持平台原编码（中文 Windows 上是 cp936/GBK）是刻意的选择，见模块 docstring。
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:  # pragma: no cover - 非 TextIOWrapper
            continue
        try:
            reconfigure(errors="replace")
        except (ValueError, OSError):  # pragma: no cover - 已关闭的流
            pass


def console_encoding() -> str:
    """返回 stdout 当前编码名（供 `speedup doctor` 体检用）。"""
    return getattr(sys.stdout, "encoding", None) or "unknown"


def display_width(text: str) -> int:
    """文本在等宽终端里占的列数：CJK 全角算 2，其余算 1，组合字符算 0。"""
    total = 0
    for ch in text:
        if unicodedata.combining(ch):
            continue
        total += 2 if unicodedata.east_asian_width(ch) in _WIDE_CATEGORIES else 1
    return total


def truncate(text: str, width: int, ellipsis: str = ELLIPSIS) -> str:
    """按显示宽度截断，超长时在末尾补省略号。``width`` 含省略号本身的宽度。"""
    if display_width(text) <= width:
        return text
    if width <= 0:
        return ""
    ellipsis_width = display_width(ellipsis)
    if width <= ellipsis_width:
        return ellipsis
    budget = width - ellipsis_width
    kept: list[str] = []
    used = 0
    for ch in text:
        w = display_width(ch)
        if used + w > budget:
            break
        kept.append(ch)
        used += w
    return "".join(kept) + ellipsis


def pad(text: str, width: int) -> str:
    """右侧补空格到指定显示宽度；已经超宽则原样返回。"""
    return text + " " * max(width - display_width(text), 0)


def render_table(
    headers: Sequence[str],
    rows: Sequence[Sequence[Any]],
    max_widths: Sequence[int | None] | None = None,
    gaps: str = "  ",
    row_prefixes: Sequence[str] | None = None,
    indent: str | None = None,
) -> str:
    """把二维数据渲染成等宽对齐的文本表格。

    :param headers: 表头文本
    :param rows: 数据行；``None`` 会渲染成空串
    :param max_widths: 每列最大显示宽度（``None`` 表示不限制），超长截断
    :param gaps: 列间分隔符（不参与宽度统计）
    :param row_prefixes: 每行前缀，例如 ``"! "``；**不参与宽度统计**，便于标注
    :param indent: 表头与分隔线的缩进。默认自动取 ``row_prefixes`` 里最宽的那个，
        这样表头不会和数据行错位
    :raises ValueError: 表头列数与某行不一致，或 ``max_widths`` / ``row_prefixes`` 长度不匹配
    """
    header_cells = [str(h) for h in headers]
    column_count = len(header_cells)
    body = [[("" if cell is None else str(cell)) for cell in row] for row in rows]

    for row in body:
        if len(row) != column_count:
            raise ValueError(
                f"表格列数不一致：表头 {column_count} 列，数据行 {len(row)} 列 -> {row!r}"
            )

    if max_widths is None:
        caps: list[int | None] = [None] * column_count
    else:
        if len(max_widths) != column_count:
            raise ValueError(
                f"max_widths 长度 {len(max_widths)} 与列数 {column_count} 不一致"
            )
        caps = list(max_widths)

    widths: list[int] = []
    for index, header in enumerate(header_cells):
        width = display_width(header)
        for row in body:
            width = max(width, display_width(row[index]))
        cap = caps[index]
        if cap is not None:
            width = min(width, cap)
        widths.append(width)

    if row_prefixes is not None and len(row_prefixes) != len(body):
        raise ValueError(
            f"row_prefixes 长度 {len(row_prefixes)} 与数据行数 {len(body)} 不一致"
        )

    def join(cells: Sequence[str]) -> str:
        return gaps.join(pad(truncate(cell, w), w) for cell, w in zip(cells, widths)).rstrip()

    if indent is None:
        indent = "" if not row_prefixes else " " * max(display_width(p) for p in row_prefixes)

    lines = [
        indent + join(header_cells),
        indent + "-" * (sum(widths) + display_width(gaps) * (column_count - 1)),
    ]
    for index, row in enumerate(body):
        prefix = row_prefixes[index] if row_prefixes is not None else ""
        lines.append(prefix + join(row))
    return "\n".join(lines)


def print_table(*args: Any, **kwargs: Any) -> None:
    """同 :func:`render_table`，但直接打印到 stdout。"""
    print(render_table(*args, **kwargs))


# ------------------------------------------------------------------ 分隔线与标题


def rule(char: str = "=", width: int = RULE_WIDTH) -> None:
    print(char * width)


def heading(title: str, char: str = "=", width: int = RULE_WIDTH) -> None:
    rule(char, width)
    print(title)
    rule(char, width)


def section(title: str, char: str = "-", width: int = RULE_WIDTH) -> None:
    rule(char, width)
    print(title)
    rule(char, width)


# ------------------------------------------------------------------ 状态提示


def ok(message: str) -> None:
    print(f"[OK]   {message}")


def warn(message: str) -> None:
    print(f"[警告] {message}")


def fail(message: str) -> None:
    print(f"[失败] {message}")


def info(message: str) -> None:
    print(f"[信息] {message}")
