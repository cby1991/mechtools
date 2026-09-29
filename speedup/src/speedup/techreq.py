"""图纸技术要求栏提取：读出「表面处理 / 热处理」要求。

为什么要这个模块
----------------
采购清单要带给加工厂报价，图纸的技术要求栏里往往写着「表面本色硬质阳极氧化」
这种必须单独报价的工序。原来只读标题栏的材料，这些信息全丢了。

分层设计（照 ``material.py`` 的约定）
--------------------------------------
* :func:`parse_craft` —— 纯函数，一条条目文本 → 备注文本
* :func:`find_tech_title` / :func:`extract_items` / :func:`pick_craft_item`
  —— 纯函数，只吃 ``(文本, x0, top)`` 序列，不碰磁盘
* :func:`find_craft_in_pdf` —— 唯一需要 pdfplumber 的一层

为什么必须按**坐标**切条目，不能按行顺序
------------------------------------------
技术要求栏和图框标题栏的纵向范围是重叠的，``extract_text()`` 吐出来的顺序是
交错的垃圾：

    技术要求： 零 件 代 号 1、边角去毛刺，倒钝； … 借(通)用件登记 …

但技术要求条目和标题**左边界完全相同**（实测 x0 一模一样），而标题栏文字贴在
页面左边缘。所以「``top`` 在标题下方」+「``x0`` 与标题对齐」两个条件就能干净切开。
实测这个切法在 139 份真实图纸上零污染。

措辞为什么不规范
----------------
同一件事在实测里有七八种写法：``表面阳极氧化为本色`` / ``表面喷砂阳极氧化为本色``
/ ``表面本色硬质阳极氧化`` / ``表面本色氧化处理`` / ``表面本色氧化`` /
``6061铝合金阳极氧化，表面本色，膜层厚度不低于5um`` / ``调质处理`` /
``表面处理参考CMF``。所以规则不是「解析出结构化字段」，而是
「**挑出工艺词 + 颜色词，按原文顺序拼回去，再补上膜厚**」——
宁可多带几个字，也别把「黑色 / 本色」这种外观验收项丢掉（报错了整批返工）。
"""

from __future__ import annotations

import os
import re
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from typing import Any

from .config import (
    CMF_MARKER,
    COLOR_WORDS,
    CRAFT_WORDS,
    TECH_ALIGN_TOLERANCE,
    TECH_TITLE_KEYWORDS,
)
from .pdfpage import words_in_page

__all__ = [
    "TechTitle",
    "CraftLookup",
    "strip_item_no",
    "parse_craft",
    "find_tech_title",
    "extract_items",
    "pick_craft_item",
    "find_craft_in_words",
    "find_craft_in_pdf",
]

#: 条目编号 —— **严格**版：只认顿号和全角点。
#:
#: 为什么不用半角点：尺寸标注 ``0.00`` 也长得像 ``数字 + 点``，实测有一份图纸
#: 的 ``0.00`` 就混进过候选区。真实图纸的条目编号清一色是顿号，收紧一点更安全。
_ITEM_NO_STRICT = re.compile(r"^\s*\d+\s*[、．]")

#: 条目编号 —— **宽松**版：剥编号时用，半角点也认（``strip_item_no`` 的输入
#: 已经是确定的条目，此时宽容不会误伤）。
_ITEM_NO_LOOSE = re.compile(r"^\s*\d+\s*[、.．]")

#: 条目收尾标点，剥编号时一并去掉
_TRAILING = "；;。.，,、 \t"

#: 膜层厚度：``膜层厚度不低于5um`` / ``膜厚≥8μm`` / ``膜厚 10 微米``
_THICKNESS = re.compile(
    r"膜[层]?厚[度]?[^0-9]{0,8}?(\d+(?:\.\d+)?)\s*(?:u\s*m|μm|µ\s*m|微米)",
    re.I,
)

#: 膜厚的比较词（决定输出带不带 ``≥``）
_GE_WORDS = re.compile(r"(不低于|不小于|至少|最少|大于等于|≥|>=)")


@dataclass(frozen=True, slots=True)
class TechTitle:
    """技术要求栏的标题词：文本 + 左边界 ``x0`` + 纵向位置 ``top``。"""

    text: str
    x0: float
    top: float


@dataclass(frozen=True, slots=True)
class CraftLookup:
    """一次工艺提取的结果。

    ``text`` 是要写进备注的文本，``raw`` 是命中的原始条目（留作追溯），
    ``error`` 非空表示「没读成」而不是「图纸没写」—— 两者必须区分开。
    """

    text: str | None = None
    raw: str | None = None
    where: str | None = None
    error: str | None = None

    @property
    def found(self) -> bool:
        return self.text is not None


def _as_float(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _fields(word: Any) -> tuple[str, float, float]:
    """从词对象里取出 ``(文本, x0, top)``。

    兼容两种形式：``extract_words()`` 返回的 dict，以及 ``(文本, x0, top)``
    元组 —— 后者是测试里手写用例用的，不依赖 pdfplumber。
    """
    if isinstance(word, Mapping):
        text = word.get("text", "")
        x0 = word.get("x0", 0.0)
        top = word.get("top", 0.0)
    elif isinstance(word, (tuple, list)) and len(word) >= 3:
        text, x0, top = word[0], word[1], word[2]
    elif isinstance(word, (tuple, list)) and word:
        text, x0, top = word[0], 0.0, 0.0
    else:
        text, x0, top = word, 0.0, 0.0
    return ("" if text is None else str(text)), _as_float(x0), _as_float(top)


def strip_item_no(text: str) -> str:
    """剥掉条目编号（``1、`` ``2.`` ``3．``）和收尾标点。"""
    if not text:
        return ""
    return _ITEM_NO_LOOSE.sub("", text).strip().rstrip(_TRAILING)


def _match_words(text: str, *word_groups: Iterable[str]) -> list[tuple[int, str]]:
    """在 ``text`` 里做「组间优先、组内最长优先」的不重叠匹配。

    :returns: ``[(起始位置, 词), ...]``，按位置升序。

    两个关键点：

    * **组间优先**：先把工艺词占的字符记下来，颜色词只能匹配**剩下**的位置。
      所以 ``本色氧化`` 吃掉「本色」之后，颜色词表里的「本色」不会再匹配一次
      （否则会得到「本色氧化本色」）。
    * **组内最长优先**：``硬质阳极氧化`` 必须赢过 ``阳极氧化``。按长度降序试，
      再靠 ``used`` 标记防止短词回头咬一口。
    """
    used = [False] * len(text)
    hits: list[tuple[int, str]] = []
    for group in word_groups:
        for word in sorted(set(group), key=len, reverse=True):
            start = 0
            while True:
                index = text.find(word, start)
                if index < 0:
                    break
                end = index + len(word)
                if not any(used[index:end]):
                    hits.append((index, word))
                    for position in range(index, end):
                        used[position] = True
                start = index + 1
    hits.sort()
    return hits


def _extract_thickness(text: str) -> str | None:
    """抽出膜厚要求，归一成 ``≥5um`` 这种形式；没有就返回 ``None``。"""
    match = _THICKNESS.search(text)
    if not match:
        return None
    # 比较词（「不低于」）夹在「膜层厚度」和数字中间，所以要在**匹配段内部**找，
    # 而不是它前面 —— 前面找到的「至少」可能是上一句话的。
    sign = "≥" if _GE_WORDS.search(match.group(0)) else ""
    return f"{sign}{match.group(1)}um"


def parse_craft(text: str) -> str | None:
    """把一条技术要求解析成备注文本；**不是工艺条目就返回 ``None``**。

    规则：

    * 命中 ``CMF`` 的**原样保留**（图纸没写具体工艺，只指向 CMF 文件，
      这条信息必须带给采购，不能当「没找到」丢掉）
    * 其余取「工艺词 + 颜色词」，**按它们在原文里的先后顺序**拼接
      —— 这样 ``表面本色硬质阳极氧化`` 会得到 ``本色硬质阳极氧化``，
      语序和图纸一致，人一眼能对上
    * 带膜厚的追加 ``·膜厚≥5um``

    :param text: 一条条目原文，可带编号，如 ``5、表面喷砂阳极氧化为本色；``
    """
    core = strip_item_no(text)
    if not core:
        return None

    if CMF_MARKER in core.upper():
        # 只把 cmf 统一成 CMF，其余原样
        return re.sub(r"(?i)cmf", CMF_MARKER, core)

    spans = _match_words(core, CRAFT_WORDS, COLOR_WORDS)
    if not spans:
        return None

    result = "".join(word for _, word in spans)
    thickness = _extract_thickness(core)
    if thickness:
        result += f"·膜厚{thickness}"
    return result


def find_tech_title(words: Iterable[Any]) -> TechTitle | None:
    """找出技术要求栏的标题词；找不到返回 ``None``。

    实测同一批图纸里「技术要求：」和「设计要求：」都在用，所以标题是一组关键词
    （见 :data:`speedup.config.TECH_TITLE_KEYWORDS`），只认一种会整批漏掉。
    """
    for word in words:
        text, x0, top = _fields(word)
        if any(keyword in text for keyword in TECH_TITLE_KEYWORDS):
            return TechTitle(text=text, x0=x0, top=top)
    return None


def extract_items(
    words: Iterable[Any],
    title: TechTitle,
    *,
    align_tolerance: float = TECH_ALIGN_TOLERANCE,
) -> list[str]:
    """取出标题下方、与标题左对齐的**条目**文本，按纵向顺序返回。

    :param words: ``{"text": str, "x0": float, "top": float}`` 序列，
        或 ``(文本, x0, top)`` 元组序列
    :param title: :func:`find_tech_title` 的结果
    :param align_tolerance: 横向对齐容差（pt）；太大会把图面尺寸数字也收进来
    """
    picked: list[tuple[float, str]] = []
    for word in words:
        text, x0, top = _fields(word)
        if top <= title.top:
            continue
        if abs(x0 - title.x0) > align_tolerance:
            continue
        if not _ITEM_NO_STRICT.match(text):
            continue
        picked.append((top, text))
    picked.sort(key=lambda item: item[0])
    return [text for _, text in picked]


def pick_craft_item(items: Iterable[str]) -> str | None:
    """从条目里挑出「讲工艺」的那一条（**取第一条命中的**）。

    实测每张图纸最多只有一条工艺要求。真出现多条时取第一条，
    好过把两条拼起来 —— 拼错了没人看得出来。
    """
    for item in items:
        if parse_craft(item) is not None:
            return item
    return None


def find_craft_in_words(
    words: Iterable[Any],
    *,
    align_tolerance: float = TECH_ALIGN_TOLERANCE,
) -> CraftLookup:
    """从词序列里提取工艺要求（**纯函数，不碰磁盘**）。

    与 :func:`find_craft_in_pdf` 的区别只在「谁来读 PDF」：编排层如果已经为了
    别的事（比如识别材料）读过这一页，就直接调这个，省一次 ``pdfplumber.open``。

    :param words: ``{"text": str, "x0": float, "top": float}`` 序列，
        或 ``(文本, x0, top)`` 元组序列
    """
    # 先物化成 list：下面要遍历两次（找标题 + 取条目），传进来的可能是生成器
    tokens = list(words)

    title = find_tech_title(tokens)
    if title is None:
        return CraftLookup()

    raw = pick_craft_item(
        extract_items(tokens, title, align_tolerance=align_tolerance)
    )
    if raw is None:
        return CraftLookup(where=title.text)
    return CraftLookup(text=parse_craft(raw), raw=raw, where=title.text)


def find_craft_in_pdf(
    pdf_path: str | os.PathLike[str],
    *,
    align_tolerance: float = TECH_ALIGN_TOLERANCE,
) -> CraftLookup:
    """读取 PDF 第一页，从技术要求栏提取工艺要求。

    读不了（没装 pdfplumber / 文件损坏 / 加密）时返回带 ``error`` 的结果，
    **不抛异常** —— 与 :func:`speedup.material.find_material_in_pdf` 保持一致，
    调用方只需看 :attr:`CraftLookup.found`。
    """
    try:
        import pdfplumber
    except ImportError:
        return CraftLookup(error="未安装 pdfplumber")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 只取页面内的词 —— 页面外可能藏着另一套「技术要求」（见 pdfpage）
            words = words_in_page(pdf.pages[0])
    except Exception as exc:  # 加密 / 损坏 / 权限
        return CraftLookup(error=f"读取失败({exc})")

    return find_craft_in_words(words, align_tolerance=align_tolerance)
