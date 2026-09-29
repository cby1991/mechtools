"""技术要求栏提取测试。

样本全部取自真实图纸里出现过的措辞 —— 这批图纸的写法非常不规范
（同一件事有七八种写法），所以每个参数化用例都是一条真实样本。
"""

from __future__ import annotations

import pytest

from speedup.config import TECH_ALIGN_TOLERANCE
from speedup.techreq import (
    CraftLookup,
    TechTitle,
    extract_items,
    find_craft_in_pdf,
    find_craft_in_words,
    find_tech_title,
    parse_craft,
    pick_craft_item,
    strip_item_no,
)


class TestStripItemNo:
    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("1、边角去毛刺，倒钝；", "边角去毛刺，倒钝"),
            ("5、表面阳极氧化为本色；", "表面阳极氧化为本色"),
            (
                "6、6061铝合金阳极氧化,表面黑色处理，膜层厚度不低于5um；",
                "6061铝合金阳极氧化,表面黑色处理，膜层厚度不低于5um",
            ),
            ("2. 表面本色喷砂", "表面本色喷砂"),
            ("3．未注倒角尺寸为C0.5；", "未注倒角尺寸为C0.5"),
            ("7、未注尺寸详见3D模型。", "未注尺寸详见3D模型"),
            ("表面电解抛光", "表面电解抛光"),
            ("", ""),
        ],
    )
    def test_strips_number_and_trailing_punctuation(self, raw, expected):
        assert strip_item_no(raw) == expected


class TestParseCraft:
    """实测措辞 → 备注文本。

    输出里的词按它们在原文中出现的先后排列（最长匹配优先），
    这样「表面本色硬质阳极氧化」会得到「本色硬质阳极氧化」而不是
    「硬质阳极氧化本色」—— 语序与图纸一致，人一眼就能对上。
    """

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("5、表面阳极氧化为本色；", "阳极氧化本色"),
            ("5、表面喷砂阳极氧化为本色；", "喷砂阳极氧化本色"),
            ("2、表面本色硬质阳极氧化；", "本色硬质阳极氧化"),
            ("2、表面喷砂，本色硬质阳极氧化；", "喷砂本色硬质阳极氧化"),
            ("5、表面本色喷砂阳极氧化；", "本色喷砂阳极氧化"),
            ("4、表面阳极氧化本色；", "阳极氧化本色"),
            ("5、表面本色氧化处理；", "本色氧化"),
            ("5、表面本色氧化", "本色氧化"),
            ("5、外观件，表面本色喷砂；", "本色喷砂"),
            ("5、外观件，表面电镀处理。", "电镀"),
            ("5、调质处理；", "调质"),
        ],
    )
    def test_common_wordings(self, raw, expected):
        assert parse_craft(raw) == expected

    def test_colour_inside_craft_word_is_not_repeated(self):
        """``本色氧化`` 本身就是一个工艺词，不该再拼一个 ``本色`` 进去。"""
        assert parse_craft("5、表面本色氧化处理；") == "本色氧化"

    def test_longer_craft_word_wins(self):
        """``硬质阳极氧化`` 必须优先于 ``阳极氧化``，否则会丢掉「硬质」。"""
        assert parse_craft("2、表面本色硬质阳极氧化；") == "本色硬质阳极氧化"


class TestParseCraftThickness:
    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            (
                "6、6061铝合金阳极氧化,表面黑色处理，膜层厚度不低于5um；",
                "阳极氧化黑色·膜厚≥5um",
            ),
            (
                "6、6061铝合金阳极氧化，表面本色，膜层厚度不低于5um；",
                "阳极氧化本色·膜厚≥5um",
            ),
            (
                "6、6061铝合金阳极氧化，本色处理，膜层厚度不低于5um；",
                "阳极氧化本色·膜厚≥5um",
            ),
        ],
    )
    def test_thickness_is_appended(self, raw, expected):
        assert parse_craft(raw) == expected

    def test_no_thickness_no_suffix(self):
        assert "膜厚" not in parse_craft("5、表面阳极氧化为本色；")


class TestParseCraftCmf:
    """图纸只说「参考 CMF」，没写具体工艺。原样保留，别丢信息。"""

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("5、表面处理参考CMF", "表面处理参考CMF"),
            ("5、表面处理参考CMF文件；", "表面处理参考CMF文件"),
            ("5、参考CMF文件。", "参考CMF文件"),
            ("6、表面处理参考cmf；", "表面处理参考CMF"),
        ],
    )
    def test_cmf_is_kept_verbatim(self, raw, expected):
        assert parse_craft(raw) == expected


class TestParseCraftReturnsNone:
    """非工艺条目必须返回 None —— 否则备注列会被「未注公差」这类话填满。"""

    @pytest.mark.parametrize(
        "raw",
        [
            "1、边角去毛刺，倒钝；",
            "2、未注公差尺寸为±0.1mm",
            "3、未注倒角尺寸为C0.5；",
            "4、未注尺寸参考三维模型；",
            "7、未注尺寸详见3D模型。",
            "4、未注尺寸公差按GB/T1804-m级执行；",
            "5、未注形位公差按GB/T1184-K执行；",
            "3、零件加工完成后无毛刺飞边；",
            "1、未注倒角C0.5;",
            "2、未注圆角R0.5；",
            "",
            "   ",
        ],
    )
    def test_not_a_craft(self, raw):
        assert parse_craft(raw) is None


class TestFindTechTitle:
    def test_finds_technical_requirement(self):
        words = [{"text": "技术要求：", "x0": 413.3, "top": 502.4}]
        assert find_tech_title(words) == TechTitle(
            text="技术要求：", x0=413.3, top=502.4
        )

    def test_finds_design_requirement(self):
        """实测有 6 份图纸用的是「设计要求：」而不是「技术要求：」。"""
        title = find_tech_title([{"text": "设计要求：", "x0": 60.0, "top": 700.0}])
        assert title is not None
        assert title.text == "设计要求："

    def test_returns_none_when_absent(self):
        assert find_tech_title([{"text": "零件代号", "x0": 11.0, "top": 533.0}]) is None

    def test_empty_words(self):
        assert find_tech_title([]) is None

    def test_accepts_tuple_form(self):
        title = find_tech_title([("技术要求", 400.0, 500.0)])
        assert title is not None
        assert title.x0 == 400.0


class TestExtractItems:
    """条目与标题左对齐；图框标题栏文字贴在页面左边缘，必须被排除。"""

    TITLE = TechTitle(text="技术要求：", x0=413.3, top=502.4)

    WORDS = [
        {"text": "技术要求：", "x0": 413.3, "top": 502.4},
        {"text": "6.02", "x0": 169.4, "top": 529.9},  # 图面尺寸，横向差得远
        {"text": "1、边角去毛刺，倒钝；", "x0": 413.3, "top": 530.5},
        {"text": "零", "x0": 11.2, "top": 533.5},  # 标题栏
        {"text": "件", "x0": 24.0, "top": 533.5},  # 标题栏
        {"text": "2、未注公差尺寸为±0.1mm", "x0": 413.3, "top": 552.7},
        {"text": "借(通)用件登记", "x0": 6.8, "top": 570.7},  # 标题栏
        {"text": "3、未注倒角尺寸为C0.5；", "x0": 413.3, "top": 575.0},
        {"text": "4、未注尺寸参考三维模型；", "x0": 413.3, "top": 597.2},
        {"text": "描图", "x0": 26.8, "top": 607.4},  # 标题栏
        {"text": "5、表面阳极氧化为本色；", "x0": 413.3, "top": 619.5},
    ]

    def test_extracts_all_items_in_reading_order(self):
        assert extract_items(self.WORDS, self.TITLE) == [
            "1、边角去毛刺，倒钝；",
            "2、未注公差尺寸为±0.1mm",
            "3、未注倒角尺寸为C0.5；",
            "4、未注尺寸参考三维模型；",
            "5、表面阳极氧化为本色；",
        ]

    def test_title_bar_words_are_excluded(self):
        joined = " ".join(extract_items(self.WORDS, self.TITLE))
        assert "零" not in joined
        assert "借(通)用件登记" not in joined
        assert "描图" not in joined

    def test_words_above_title_are_excluded(self):
        words = [
            {"text": "技术要求：", "x0": 413.3, "top": 502.4},
            {"text": "6.02", "x0": 413.3, "top": 495.0},  # 标题上方
        ]
        assert extract_items(words, self.TITLE) == []

    def test_word_off_the_left_edge_is_excluded(self):
        """回归：实测有一份图纸混进过 ``0.00``（x0 只差 7pt，比容差大一点）。"""
        offset = TECH_ALIGN_TOLERANCE + 1.0
        words = [
            {"text": "技术要求：", "x0": 413.3, "top": 502.4},
            {"text": "0.00", "x0": 413.3 - offset, "top": 600.0},
        ]
        assert extract_items(words, self.TITLE) == []

    def test_decimal_dimension_is_not_an_item(self):
        """回归：``0.00`` 和条目编号只差一个标点（半角点 vs 顿号）。

        实测有一份图纸的 ``0.00`` 曾混进候选区，所以条目编号收紧到只认顿号 /
        全角点 —— 对齐容差是最后一道防线，不该是唯一一道。
        """
        words = [
            {"text": "技术要求：", "x0": 413.3, "top": 502.4},
            {"text": "0.00", "x0": 413.3, "top": 600.0},  # x0 完全对齐
        ]
        assert extract_items(words, self.TITLE) == []

    def test_no_items(self):
        words = [{"text": "技术要求：", "x0": 413.3, "top": 502.4}]
        assert extract_items(words, self.TITLE) == []


class TestPickCraftItem:
    def test_picks_the_craft_line(self):
        items = ["1、边角去毛刺，倒钝；", "5、表面阳极氧化为本色；"]
        assert pick_craft_item(items) == "5、表面阳极氧化为本色；"

    def test_returns_none_when_no_craft_line(self):
        assert pick_craft_item(["1、边角去毛刺，倒钝；"]) is None

    def test_empty(self):
        assert pick_craft_item([]) is None

    def test_picks_first_when_several(self):
        items = ["5、表面阳极氧化为本色；", "6、表面喷砂；"]
        assert pick_craft_item(items) == "5、表面阳极氧化为本色；"


class TestFindCraftInWords:
    """纯函数入口：只吃词序列，不碰磁盘。编排层可以一次读 PDF 再分发到这里。"""

    WORDS = [
        {"text": "技术要求：", "x0": 413.3, "top": 502.4},
        {"text": "1、边角去毛刺，倒钝；", "x0": 413.3, "top": 530.5},
        {"text": "零", "x0": 11.2, "top": 533.5},  # 标题栏
        {"text": "2、未注公差尺寸为±0.1mm", "x0": 413.3, "top": 552.7},
        {"text": "5、表面喷砂阳极氧化为本色；", "x0": 413.3, "top": 619.5},
    ]

    def test_extracts_craft(self):
        result = find_craft_in_words(self.WORDS)
        assert result.text == "喷砂阳极氧化本色"
        assert result.raw == "5、表面喷砂阳极氧化为本色；"
        assert result.where == "技术要求："
        assert result.error is None

    def test_no_title_at_all(self):
        result = find_craft_in_words([{"text": "零件代号", "x0": 11.0, "top": 533.0}])
        assert result.text is None
        assert result.where is None
        assert result.error is None

    def test_title_but_no_craft_line(self):
        words = [
            {"text": "技术要求：", "x0": 413.3, "top": 502.4},
            {"text": "1、边角去毛刺，倒钝；", "x0": 413.3, "top": 530.5},
        ]
        result = find_craft_in_words(words)
        assert result.text is None
        assert result.where == "技术要求："
        assert result.error is None


class TestFindCraftInPdf:
    def test_missing_file_reports_error_without_raising(self, tmp_path):
        result = find_craft_in_pdf(tmp_path / "没有这个文件.pdf")
        assert result.text is None
        assert result.error
        assert result.found is False

    def test_garbage_file_reports_error(self, tmp_path):
        bogus = tmp_path / "假的.pdf"
        bogus.write_bytes(b"not a pdf at all")
        result = find_craft_in_pdf(bogus)
        assert result.text is None
        assert result.error
        assert result.found is False


class TestCraftLookup:
    def test_found_flag(self):
        assert CraftLookup(text="阳极氧化本色").found is True
        assert CraftLookup().found is False
