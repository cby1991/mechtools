"""材料识别测试：规则、标题栏优先级、PDF 读取的失败路径。"""

from __future__ import annotations

import pytest

from speedup.material import (
    MaterialHit,
    MaterialLookup,
    detect_material,
    find_material_in_pdf,
    find_material_in_words,
    index_pdf_by_stem,
    lookup_material_for,
    strip_prefix,
)


class TestDetectMaterial:
    @pytest.mark.parametrize(
        ("text", "expected"),
        [
            ("7075-T6", "7075-T6"),
            ("6061-T6", "6061-T6"),
            ("2A12-T4", "2A12-T4"),
            ("3Cr13", "3Cr13"),
            ("2Cr13", "2Cr13"),
            ("Q235", "Q235"),
            ("45钢", "45钢"),
            ("TC4", "TC4"),
            ("TA2", "TA2"),
            ("POM", "POM"),
            ("ABS", "ABS"),
            ("PC", "PC"),
            ("PA66", "PA66"),
            ("PEEK", "PEEK"),
            ("铝合金", "铝合金"),
            ("不锈钢", "不锈钢"),
            ("钛合金", "钛合金"),
            ("紫铜", "紫铜"),
            ("黄铜", "黄铜"),
        ],
    )
    def test_strong_patterns_match_anywhere(self, text, expected):
        assert detect_material(text) == expected

    @pytest.mark.parametrize("text", ["304", "316L", "H62", "201", "430", "2A12", "5A06"])
    def test_weak_patterns_are_rejected_by_default(self, text):
        """纯数字牌号（不锈钢 / 铜合金 / 铝合金）和尺寸、图号长得一样，默认不认。"""
        assert detect_material(text) is None

    @pytest.mark.parametrize("text", ["304", "316L", "H62", "201", "430", "2A12", "5A06"])
    def test_weak_patterns_allowed_on_demand(self, text):
        assert detect_material(text, allow_weak=True) == text

    def test_aluminium_grade_with_temper_is_strong(self):
        """``2A12-T4`` 带状态后缀，形状够独特，全文命中即可信。

        回归测试：旧版规则只认 3~4 位数字 + 破折号，把 ``2A12-T4`` 漏掉了
        （源码注释声称支持，实际匹配不上）。
        """
        assert detect_material("2A12-T4") == "2A12-T4"
        assert detect_material("5A06-H112") == "5A06-H112"

    @pytest.mark.parametrize(
        "text",
        ["材料：7075-T6", "材料:7075-T6", "材质：Q235", "材质:Q235"],
    )
    def test_prefix_is_stripped(self, text):
        assert detect_material(text) in {"7075-T6", "Q235"}

    def test_prefix_helpers(self):
        assert strip_prefix("材料：7075-T6") == "7075-T6"
        assert strip_prefix("7075-T6") == "7075-T6"

    def test_plain_dimension_is_not_material(self):
        assert detect_material("1260") is None
        assert detect_material("R5") is None

    def test_empty_text(self):
        assert detect_material("") is None

    def test_returns_first_match(self):
        assert detect_material("螺母 7075-T6 垫片 Q235") == "7075-T6"


class TestFindMaterialInWords:
    """``words`` 一般由 pdfplumber 的 ``extract_words()`` 给出。"""

    PAGE_HEIGHT = 1000.0

    def test_finds_in_header_zone(self):
        words = [{"text": "标题", "top": 100}, {"text": "7075-T6", "top": 900}]
        hit = find_material_in_words(words, self.PAGE_HEIGHT)
        assert hit == MaterialHit(material="7075-T6", where="标题栏")

    def test_finds_in_body_with_strong_pattern(self):
        words = [{"text": "Q235", "top": 100}]
        hit = find_material_in_words(words, self.PAGE_HEIGHT)
        assert hit == MaterialHit(material="Q235", where="全文")

    def test_weak_pattern_only_recognised_in_header_zone(self):
        """正文里的纯数字牌号不认（避免把尺寸 304 当不锈钢）。"""
        assert find_material_in_words([{"text": "304", "top": 100}], self.PAGE_HEIGHT) is None
        hit = find_material_in_words([{"text": "304", "top": 900}], self.PAGE_HEIGHT)
        assert hit == MaterialHit(material="304", where="标题栏")

    def test_header_zone_wins_over_body(self):
        words = [
            {"text": "Q235", "top": 100},  # 正文
            {"text": "7075-T6", "top": 950},  # 标题栏
        ]
        hit = find_material_in_words(words, self.PAGE_HEIGHT)
        assert hit is not None
        assert hit.material == "7075-T6"
        assert hit.where == "标题栏"

    def test_accepts_tuple_form(self):
        hit = find_material_in_words([("材质:Q235", 900)], self.PAGE_HEIGHT)
        assert hit is not None
        assert hit.material == "Q235"

    def test_no_material_found(self):
        assert find_material_in_words([{"text": "M8", "top": 900}], self.PAGE_HEIGHT) is None

    def test_custom_header_ratio(self):
        words = [{"text": "304", "top": 500}]
        assert find_material_in_words(words, 1000.0, header_ratio=0.4) is not None
        assert find_material_in_words(words, 1000.0, header_ratio=0.9) is None


class TestFindMaterialInPdf:
    def test_missing_file_reports_error_without_raising(self, tmp_path):
        result = find_material_in_pdf(tmp_path / "没有这个文件.pdf")
        assert result.material is None
        assert result.error
        assert result.found is False

    def test_garbage_file_reports_error(self, tmp_path):
        bogus = tmp_path / "假的.pdf"
        bogus.write_bytes(b"not a pdf at all")
        result = find_material_in_pdf(bogus)
        assert result.material is None
        assert result.error


class TestLookupHelpers:
    def test_index_uses_folder_and_stem(self, tmp_path, make_files):
        make_files(tmp_path, ["a.pdf", "b.pdf", "c.STEP"])
        make_files(tmp_path / "sub", ["a.pdf"])
        index = index_pdf_by_stem(tmp_path)

        assert (str(tmp_path).lower(), "a") in index
        assert (str(tmp_path / "sub").lower(), "a") in index
        assert (str(tmp_path).lower(), "c") not in index

    def test_lookup_requires_same_folder(self, tmp_path, make_files):
        """不同文件夹里的同名 PDF 不该被当成同一个零件的图纸。"""
        (model,) = make_files(tmp_path, ["关节支撑板.STEP"])
        make_files(tmp_path / "别处", ["关节支撑板.pdf"])

        index = index_pdf_by_stem(tmp_path)
        result, pdf_path = lookup_material_for(model, index)
        assert pdf_path is None
        assert result.material is None
        assert result.error is None

    def test_lookup_finds_same_folder_pdf(self, tmp_path, make_files):
        (model,) = make_files(tmp_path, ["垫片.STEP"])
        (pdf,) = make_files(tmp_path, ["垫片.pdf"])
        index = index_pdf_by_stem(tmp_path)
        _result, pdf_path = lookup_material_for(model, index)
        assert pdf_path == str(pdf)

    def test_lookup_case_insensitive_extension(self, tmp_path, make_files):
        (model,) = make_files(tmp_path, ["垫片.STEP"])
        make_files(tmp_path, ["垫片.PDF"])
        index = index_pdf_by_stem(tmp_path)
        _result, pdf_path = lookup_material_for(model, index)
        assert pdf_path is not None


class TestMaterialLookup:
    def test_found_flag(self):
        assert MaterialLookup(material="Q235", where="标题栏").found is True
        assert MaterialLookup().found is False
