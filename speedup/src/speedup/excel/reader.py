"""读取 xlsx，并从中抽取「名称」列的零件名。

为什么有两个解析器
------------------
openpyxl 对大多数文件没问题，但遇到「含特殊视图属性 / 非标准生成器产出」的
xlsx 会直接抛异常（现实里经常遇到别的系统导出的表）。所以保留一条**零依赖**
的回退路径：直接把 xlsx 当 zip 解开，用 :mod:`xml.etree` 读
``sharedStrings.xml`` 与 ``sheetN.xml``。

``.xls``（Excel 97-2003 的二进制格式）**不支持**，请先另存为 ``.xlsx``。

一份工作簿只认一张表
--------------------
现实的补料清单常常不止一张表带「名称」列 —— 后面还跟着「备料参考」
「备料明细」之类的附页。**只读其中一张**，选法见
:func:`extract_names_from_sheets`（优先认名为「补料清单」的表）。

早期实现是把所有表的名**合并**，于是参考页的零件也被当成待补料，
白拷了一堆文件出去。这个坑记在 ``AGENTS.md`` 的 P12。

分层设计
--------
:func:`extract_names_from_sheets` 是纯函数，输入 ``{工作表名: 二维字符串}``，
不碰磁盘 —— 单测直接喂字典即可。
"""

from __future__ import annotations

import os
import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass

from ..config import (
    HEADER_SEARCH_ROWS,
    NAME_HEADER_KEYWORD,
    PREFERRED_INPUT_SHEET,
)
from ..errors import InputError, UnsupportedFormatError

__all__ = [
    "Sheets",
    "NameExtraction",
    "column_letters_to_index",
    "read_xlsx_with_openpyxl",
    "read_xlsx_builtin",
    "read_workbook",
    "extract_names_from_sheets",
    "extract_names",
]

#: ``{工作表名: [[单元格文本, ...], ...]}``
Sheets = dict[str, list[list[str]]]

MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

_SKIP_MARKERS = {"", "none", "nan"}


@dataclass(frozen=True, slots=True)
class _Candidate:
    """一张「前 10 行里有表头」的工作表。"""

    name: str
    column: int  # 0 基
    header_row: int  # 0 基


@dataclass(frozen=True, slots=True)
class NameExtraction:
    """一次名称抽取的结果。

    为什么不是简单的 ``(names, info)`` 二元组：调用方需要**判断**这次读取是否
    有歧义（多张表都像清单），才能决定要不要警告用户。把判断依据塞进一句
    给人看的字符串里，调用方就只能靠解析文本 —— 那正是这个 bug 当初藏身的地方。

    :param names: 去重保序的零件名
    :param sheet: 实际采用的工作表；``None`` 表示一张候选表都没有
    :param column: 1 基列号（给人看的）；指定了表但没找到表头时为 ``None``
    :param rows: 该列一共读到多少行（去重**之前**的数量）
    :param skipped: 同样带表头、但没被采用的工作表
    :param requested: 工作表是调用方**明确指定**的（而不是自动挑的）
    :param header_keyword: 本次认的表头文字
    """

    names: tuple[str, ...] = ()
    sheet: str | None = None
    column: int | None = None
    rows: int = 0
    skipped: tuple[str, ...] = ()
    requested: bool = False
    header_keyword: str = NAME_HEADER_KEYWORD

    @property
    def ambiguous(self) -> bool:
        """工作簿里是否存在**多张**像清单的表（纯事实，不代表要提醒）。"""
        return bool(self.skipped)

    def summary(self) -> str:
        """一句话说明读到了什么，直接拿来打印。"""
        if self.sheet is None:
            return f"未找到名为「{self.header_keyword}」的列"
        if self.column is None:
            return f"工作表「{self.sheet}」里没有「{self.header_keyword}」列"
        return f"工作表「{self.sheet}」第 {self.column} 列，读到 {self.rows} 行"

    def warning(self) -> str | None:
        """需要提醒用户「我替你做了个选择」时给一句；不需要则返回 ``None``。

        两种情况下不提醒：

        * 只有一张候选表 —— 没什么可选的
        * 用户**明确指定**了表（``--sheet``）—— 他已经做过决定了，
          再推荐他去用 ``--sheet`` 是废话
        """
        if not self.skipped or self.requested:
            return None
        others = "、".join(f"「{name}」" for name in self.skipped)
        return (
            f"这份工作簿里有 {len(self.skipped) + 1} 张表都带「{self.header_keyword}」列，"
            f"本次只用「{self.sheet}」，忽略了{others}。"
            f"如果那些零件也该一起补料，请用 --sheet 指定要读哪张表。"
        )


def column_letters_to_index(letters: str) -> int:
    """Excel 列字母转 0 基下标：``'A' -> 0``，``'C' -> 2``，``'AA' -> 26``。"""
    index = 0
    for char in letters:
        index = index * 26 + (ord(char.upper()) - 64)
    return index - 1


def read_xlsx_with_openpyxl(path: str | os.PathLike[str]) -> Sheets:
    """用 openpyxl 读取全部工作表（优先路径）。"""
    import openpyxl

    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        sheets: Sheets = {}
        for worksheet in workbook.worksheets:
            rows: list[list[str]] = []
            for row in worksheet.iter_rows(values_only=True):
                rows.append(["" if value is None else str(value) for value in row])
            sheets[worksheet.title] = rows
        return sheets
    finally:
        workbook.close()


def read_xlsx_builtin(path: str | os.PathLike[str]) -> Sheets:
    """零依赖解析 xlsx（openpyxl 不可用或解析失败时的回退）。"""
    with zipfile.ZipFile(path) as archive:
        members = set(archive.namelist())

        shared: list[str] = []
        if "xl/sharedStrings.xml" in members:
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall(MAIN_NS + "si"):
                shared.append("".join(node.text or "" for node in item.iter(MAIN_NS + "t")))

        relationships: dict[str, str] = {}
        if "xl/_rels/workbook.xml.rels" in members:
            for rel in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels")):
                relationships[rel.get("Id", "")] = rel.get("Target", "")

        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet_nodes = workbook_root.find(MAIN_NS + "sheets")
        if sheet_nodes is None:
            return {}

        sheets: Sheets = {}
        for sheet_node in sheet_nodes:
            target = relationships.get(sheet_node.get(REL_NS + "id", ""), "").lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            if target not in members:
                continue
            sheets[sheet_node.get("name", target)] = _parse_sheet(archive.read(target), shared)
        return sheets


def _parse_sheet(payload: bytes, shared: Sequence[str]) -> list[list[str]]:
    root = ET.fromstring(payload)
    rows: list[list[str]] = []
    for row_node in root.iter(MAIN_NS + "row"):
        cells: dict[int, str] = {}
        for cell in row_node.findall(MAIN_NS + "c"):
            letters = "".join(ch for ch in (cell.get("r") or "") if ch.isalpha())
            if not letters:
                continue
            cell_type = cell.get("t")
            value_node = cell.find(MAIN_NS + "v")
            if cell_type == "s" and value_node is not None:
                try:
                    value = shared[int(value_node.text or "0")]
                except (ValueError, IndexError):
                    value = ""
            elif cell_type == "inlineStr":
                value = "".join(node.text or "" for node in cell.iter(MAIN_NS + "t"))
            elif value_node is not None:
                value = value_node.text or ""
            else:
                value = ""
            cells[column_letters_to_index(letters)] = value
        if not cells:
            rows.append([])
            continue
        width = max(cells) + 1
        rows.append([cells.get(i, "") or "" for i in range(width)])
    return rows


def read_workbook(path: str | os.PathLike[str]) -> Sheets:
    """读取工作簿，自动在 openpyxl 与内置解析器之间选择。

    :raises UnsupportedFormatError: 传入的是 ``.xls`` 老格式
    :raises FileNotFoundError: 文件不存在
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(f"找不到清单文件：{path}")
    if str(path).lower().endswith(".xls"):
        raise UnsupportedFormatError(
            "不支持旧的 .xls 格式，请在 Excel 中另存为 .xlsx 后再运行。"
        )
    try:
        return read_xlsx_with_openpyxl(path)
    except ImportError:
        return read_xlsx_builtin(path)
    except Exception:
        # openpyxl 对部分文件（特殊视图属性等）会解析失败，改走内置解析器
        return read_xlsx_builtin(path)


def _find_name_header(
    rows: Sequence[Sequence[object]],
    header_keyword: str,
) -> tuple[int, int] | None:
    """在前 ``HEADER_SEARCH_ROWS`` 行里找表头，返回 ``(0 基列号, 0 基行号)``。"""
    for row_index, row in enumerate(rows[:HEADER_SEARCH_ROWS]):
        for column_index, value in enumerate(row):
            if value is not None and str(value).strip() == header_keyword:
                return column_index, row_index
    return None


def _read_column(
    rows: Sequence[Sequence[object]],
    column: int,
    header_row: int,
) -> tuple[list[str], int]:
    """读表头下面那一列，去重保序。返回 ``(名称列表, 读到的行数)``。"""
    names: list[str] = []
    seen: set[str] = set()
    count = 0
    for row in rows[header_row + 1 :]:
        if column >= len(row):
            continue
        value = str(row[column] or "").strip()
        if value.lower() in _SKIP_MARKERS:
            continue
        count += 1
        if value not in seen:
            seen.add(value)
            names.append(value)
    return names, count


def _candidate_sheets(
    sheets: Mapping[str, Sequence[Sequence[object]]],
    header_keyword: str,
) -> list[_Candidate]:
    """挑出所有「前 10 行里有该表头」的工作表，按工作簿里的顺序。"""
    found: list[_Candidate] = []
    for sheet_name, rows in sheets.items():
        position = _find_name_header(rows, header_keyword)
        if position is None:
            continue
        found.append(_Candidate(name=sheet_name, column=position[0], header_row=position[1]))
    return found


def extract_names_from_sheets(
    sheets: Mapping[str, Sequence[Sequence[object]]],
    header_keyword: str = NAME_HEADER_KEYWORD,
    sheet: str | None = None,
) -> NameExtraction:
    """从工作簿里抽取「名称」列的零件名。

    **只认一张表。** 选表的顺序：

    1. 调用方用 ``sheet=`` 明确指定 → 就用它（不存在则抛 :class:`InputError`）
    2. 否则工作簿里有名为 :data:`PREFERRED_INPUT_SHEET`（「补料清单」）的表 → 用它
    3. 否则用**第一张**带该表头的表

    剩下那些同样带表头却没被选中的表记在 :attr:`NameExtraction.skipped` 里，
    由调用方决定要不要提醒用户 —— 它们常常是「参考」「备料」之类的附页。

    > 早期实现是把**所有**表的名**合并**。遇到「补料清单 + 备料参考」
    > 这种工作簿，参考页的零件也会被当成待补料，白拷一堆文件出去。
    > 见 `AGENTS.md` 的 P12。

    :returns: :class:`NameExtraction`；一张候选表都没有时其 ``sheet`` 为 ``None``
    :raises InputError: ``sheet`` 指定的工作表在工作簿里不存在
    """
    candidates = _candidate_sheets(sheets, header_keyword)
    requested = sheet is not None

    if requested:
        if sheet not in sheets:
            available = "、".join(sheets) or "（没有任何工作表）"
            raise InputError(f"工作簿里没有这张工作表：{sheet}。可选的有：{available}")
        chosen = next((c for c in candidates if c.name == sheet), None)
        if chosen is None:
            # 表在，但前 10 行里没有这个表头；交给调用方按「读到 0 个」处理
            return NameExtraction(
                names=(),
                sheet=sheet,
                column=None,
                rows=0,
                skipped=tuple(c.name for c in candidates),
                requested=True,
                header_keyword=header_keyword,
            )
    elif not candidates:
        return NameExtraction(
            names=(),
            sheet=None,
            column=None,
            rows=0,
            skipped=(),
            header_keyword=header_keyword,
        )
    else:
        chosen = next((c for c in candidates if c.name == PREFERRED_INPUT_SHEET), candidates[0])

    names, count = _read_column(sheets[chosen.name], chosen.column, chosen.header_row)
    return NameExtraction(
        names=tuple(names),
        sheet=chosen.name,
        column=chosen.column + 1,  # 1 基，给人看的
        rows=count,
        skipped=tuple(c.name for c in candidates if c.name != chosen.name),
        requested=requested,
        header_keyword=header_keyword,
    )


def extract_names(
    excel_path: str | os.PathLike[str],
    header_keyword: str = NAME_HEADER_KEYWORD,
    sheet: str | None = None,
) -> NameExtraction:
    """读文件 + 抽取名称。

    :raises InputError: 表格里没有「名称」列，或指定的工作表不存在
    """
    result = extract_names_from_sheets(
        read_workbook(excel_path), header_keyword, sheet=sheet
    )
    if not result.names:
        raise InputError(
            f"没有读到任何名称。{result.summary()}。"
            f"请确认表头里有「{header_keyword}」这一列。"
        )
    return result
