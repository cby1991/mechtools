"""图纸标题栏材料识别。

分层设计（为了可测）
--------------------
* :func:`detect_material` —— 纯函数，输入一段文字，输出材料牌号。全部规则在这里。
* :func:`find_material_in_words` —— 纯函数，输入「词 + 纵坐标」序列和页高，
  决定去标题栏还是全文找。
* :func:`find_material_in_pdf` —— 唯一需要 pdfplumber 的一层，负责读 PDF 取词。

这样 90% 的逻辑可以在不碰 PDF、不装 pdfplumber 的情况下用 pytest 覆盖。

识别规则
--------
分强弱两档：

* **强特征**（:data:`STRONG_PATTERNS`）—— ``7075-T6``、``3Cr13``、``Q235``、
  ``45钢``、``TC4``、``POM/ABS/PC…``、``铝合金/不锈钢…``。这些形状独特，
  全文任何位置命中都认。
* **弱特征**（:data:`WEAK_PATTERN`）—— ``304`` / ``316L`` / ``H62`` 这类
  纯数字或两字母牌号。它们和尺寸数字长得一模一样，所以**只在标题栏区域**认。

原文里这段逻辑是硬编码在一个函数里的，这里拆开并把阈值参数化。
"""

from __future__ import annotations

import os
import re
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from typing import Any

from .config import HEADER_ZONE_RATIO
from .pdfpage import words_in_page

__all__ = [
    "STRONG_PATTERNS",
    "WEAK_PATTERN",
    "PREFIXES",
    "MaterialHit",
    "MaterialLookup",
    "strip_prefix",
    "detect_material",
    "find_material_in_words",
    "find_material_in_pdf",
]

#: 强特征：形状足够独特，全文命中即可信
STRONG_PATTERNS: tuple[re.Pattern[str], ...] = (
    re.compile(r"\d{3,4}[-－][A-Za-z]\d+(?:\.\d+)?", re.I),  # 7075-T6 / 6061-T6 / 2A12-T4
    re.compile(r"\b[1-8][A-Za-z]\d{2}[-－][A-Za-z]\d+\b", re.I),  # 铝合金牌号带状态：2A12-T4
    re.compile(r"\d(?:\.\d+)?Cr\d+", re.I),  # 3Cr13 / 2Cr13
    re.compile(r"Q\d{3}", re.I),  # Q235
    re.compile(r"\d{1,3}#?钢"),  # 45钢
    re.compile(r"\bTC\d|\bTA\d", re.I),  # 钛合金 TC4
    re.compile(r"\b(POM|PA\d+|ABS|PP|PEEK|PTFE|PC|PVC)\b", re.I),
    re.compile(r"(黄铜|紫铜|不锈钢|铝合金|钛合金|铬铜|纯铁)"),
)

#: 弱特征：形状和尺寸/图号容易混淆，**只在标题栏区域**认。
#: 纯数字不锈钢牌号 ``304``、铜合金 ``H62``、铝合金牌号 ``2A12``（不带状态后缀）。
WEAK_PATTERN: re.Pattern[str] = re.compile(
    r"^(?:304|304L|316|316L|201|430|H59|H62|[1-8][A-Za-z]\d{2})$", re.I
)

#: 标题栏里常见的「材料：」前缀，识别前先剥掉
PREFIXES: tuple[str, ...] = ("材料：", "材料:", "材质：", "材质:")


@dataclass(frozen=True, slots=True)
class MaterialHit:
    """识别结果：牌号 + 命中位置（``"标题栏"`` 或 ``"全文"``）。"""

    material: str
    where: str


@dataclass(frozen=True, slots=True)
class MaterialLookup:
    """一次完整查找的结果，把「没找到」和「读失败」区分开。"""

    material: str | None = None
    where: str | None = None
    error: str | None = None

    @property
    def found(self) -> bool:
        return self.material is not None


def strip_prefix(text: str) -> str:
    """剥掉 ``材料：`` / ``材质:`` 之类前缀（含全角和半角冒号）。"""
    for prefix in PREFIXES:
        if text.startswith(prefix):
            return text[len(prefix) :]
    return text


def detect_material(text: str, *, allow_weak: bool = False) -> str | None:
    """从一段文字里抽出材料牌号；抽不到返回 ``None``。

    :param allow_weak: 是否允许弱特征（纯数字牌号）。只应在标题栏区域用 ``True``。
    """
    if not text:
        return None
    candidate = strip_prefix(text.strip())
    for pattern in STRONG_PATTERNS:
        match = pattern.search(candidate)
        if match:
            return match.group(0)
    if allow_weak and WEAK_PATTERN.match(candidate):
        return candidate
    return None


def _word_text(word: Any) -> str:
    """兼容 pdfplumber 的 dict 形式和 ``(文本, 纵坐标)`` 元组形式。"""
    if isinstance(word, Mapping):
        value = word.get("text", "")
    elif isinstance(word, (tuple, list)) and word:
        value = word[0]
    else:
        value = word
    return "" if value is None else str(value)


def _word_top(word: Any) -> float:
    if isinstance(word, Mapping):
        value = word.get("top", 0.0)
    elif isinstance(word, (tuple, list)) and len(word) > 1:
        value = word[1]
    else:
        value = 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def find_material_in_words(
    words: Iterable[Any],
    page_height: float,
    *,
    header_ratio: float = HEADER_ZONE_RATIO,
) -> MaterialHit | None:
    """在页面文字里找材料。先扫标题栏区域（允许弱特征），再全文扫（只认强特征）。

    :param words: ``{"text": str, "top": float}`` 序列，或 ``(文本, top)`` 元组序列
    :param page_height: 页面高度，用于换算标题栏区域
    :param header_ratio: 超过 ``page_height * header_ratio`` 的算标题栏区域
    """
    tokens = [(_word_text(word), _word_top(word)) for word in words]
    zone_top = page_height * header_ratio

    for text, top in tokens:
        if top > zone_top:
            hit = detect_material(text, allow_weak=True)
            if hit:
                return MaterialHit(material=hit, where="标题栏")

    for text, _top in tokens:
        hit = detect_material(text, allow_weak=False)
        if hit:
            return MaterialHit(material=hit, where="全文")

    return None


def find_material_in_pdf(
    pdf_path: str | os.PathLike[str],
    *,
    header_ratio: float = HEADER_ZONE_RATIO,
) -> MaterialLookup:
    """读取 PDF 第一页并识别材料。

    读不了（没装 pdfplumber、文件损坏、加密）时返回带 ``error`` 的结果，
    **不抛异常** —— 调用方只需看 :attr:`MaterialLookup.material` 是否为空。
    """
    try:
        import pdfplumber
    except ImportError:
        return MaterialLookup(error="未安装 pdfplumber")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            page_height = float(page.height)
            # 只取页面内的词 —— 页面外可能是被裁掉的另一张图（见 pdfpage）
            words = words_in_page(page)
    except Exception as exc:  # 加密 / 损坏 / 权限
        return MaterialLookup(error=f"读取失败({exc})")

    hit = find_material_in_words(words, page_height, header_ratio=header_ratio)
    if hit is None:
        return MaterialLookup()
    return MaterialLookup(material=hit.material, where=hit.where)


def index_pdf_by_stem(folder: str | os.PathLike[str]) -> dict[tuple[str, str], str]:
    """把文件夹（含子目录）里所有 PDF 建成 ``{(所属目录小写, 文件名小写): 完整路径}``。

    用「所在目录 + 同名主干」做键，是为了保证 ``xxx.stp`` 找到的是**同一个文件夹里**
    的 ``xxx.pdf``，而不是别处同名图纸。
    """
    index: dict[tuple[str, str], str] = {}
    for dirpath, _dirnames, filenames in os.walk(folder):
        for filename in filenames:
            stem, ext = os.path.splitext(filename)
            if ext.lower() != ".pdf":
                continue
            index.setdefault(
                (dirpath.lower(), stem.lower()), os.path.join(dirpath, filename)
            )
    return index


def lookup_material_for(
    model_path: str | os.PathLike[str],
    pdf_index: Mapping[tuple[str, str], str],
) -> tuple[MaterialLookup, str | None]:
    """给定模型路径，找同目录同名 PDF 并识别材料。

    :returns: ``(查找结果, 命中的 PDF 路径或 None)``
    """
    directory = os.path.dirname(os.path.abspath(model_path))
    stem = os.path.splitext(os.path.basename(model_path))[0]
    pdf_path = pdf_index.get((directory.lower(), stem.lower()))
    if pdf_path is None:
        return MaterialLookup(), None
    return find_material_in_pdf(pdf_path), pdf_path
