"""控制台输出工具的测试：中文宽度、截断、表格对齐。"""

from __future__ import annotations

import pytest

from speedup import console


def _char_at(line: str, offset: int) -> str:
    """返回 ``line`` 在「显示列 offset」处的字符（越界返回空格）。

    不能用 ``line[offset]``：汉字占 2 列，字符下标和显示列不是一回事。
    """
    used = 0
    for char in line:
        if used >= offset:
            return char
        used += console.display_width(char)
    return " "


class TestDisplayWidth:
    def test_ascii_counts_one(self):
        assert console.display_width("abc123") == 6

    def test_cjk_counts_two(self):
        assert console.display_width("中文") == 4

    def test_mixed(self):
        # "序1" = 2 + 1
        assert console.display_width("序1") == 3

    def test_empty(self):
        assert console.display_width("") == 0

    def test_fullwidth_punct_is_wide(self):
        # 全角括号是 F 类，占 2 列
        assert console.display_width("（") == 2

    def test_combining_mark_counts_zero(self):
        # e + 组合重音符 -> 只算 1 列
        assert console.display_width("e\u0301") == 1


class TestTruncate:
    def test_shorter_than_limit_is_untouched(self):
        assert console.truncate("abc", 10) == "abc"

    def test_exactly_at_limit_is_untouched(self):
        assert console.truncate("abc", 3) == "abc"

    def test_cjk_respects_double_width(self):
        # 3 个汉字 = 6 列，限宽 4 -> 只能留 1 个汉字 + 省略号(1 列)
        assert console.truncate("中文名", 4) == "中…"

    def test_ellipsis_included_in_budget(self):
        assert console.display_width(console.truncate("abcdefgh", 5)) == 5

    def test_zero_width_returns_empty_budget(self):
        assert console.truncate("abc", 0) == ""


class TestPad:
    def test_pads_ascii(self):
        assert console.pad("ab", 5) == "ab   "

    def test_pads_by_display_width(self):
        assert console.pad("中文", 6) == "中文  "

    def test_never_shrinks(self):
        assert console.pad("abcdef", 3) == "abcdef"


class TestRenderTable:
    def test_basic_shape(self):
        text = console.render_table(("A", "B"), [("1", "2"), ("3", "4")])
        lines = text.splitlines()
        assert len(lines) == 4  # 表头 + 分隔线 + 2 行
        assert set(lines[1]) == {"-"}

    def test_columns_line_up_in_display_columns(self):
        text = console.render_table(("编号", "名称"), [("1", "垫片"), ("22", "关节支撑板")])
        lines = text.splitlines()
        # 第一列宽 = max(4, 1, 2) = 4，列间距 2 -> 第二列从显示列 6 开始
        for line in [lines[0], *lines[2:]]:
            assert _char_at(line, 6) != " ", f"第二列没在第 6 显示列开始：{line!r}"

    def test_separator_spans_full_table_width(self):
        text = console.render_table(("编号", "名称"), [("1", "垫片")])
        lines = text.splitlines()
        assert set(lines[1]) == {"-"}
        assert console.display_width(lines[1]) == console.display_width(lines[0].rstrip())

    def test_truncates_to_max_width(self):
        text = console.render_table(("名称",), [("关节支撑板",)], max_widths=(6,))
        body = text.splitlines()[2]
        assert console.display_width(body.rstrip()) <= 6
        assert body.rstrip().endswith("…")

    def test_none_becomes_empty(self):
        text = console.render_table(("A", "B"), [(None, "x")])
        assert "x" in text

    def test_row_prefix_does_not_break_alignment(self):
        text = console.render_table(
            ("编号", "名称"),
            [("1", "垫片"), ("2", "法兰盘")],
            row_prefixes=["  ", "X "],
        )
        lines = text.splitlines()
        # 表头前应自动补 2 个空格，与 "X " 的宽度对齐
        assert lines[0].startswith("  编号")
        assert lines[1].startswith("  ---")
        assert lines[2].startswith("  1")
        assert lines[3].startswith("X 2")

    def test_explicit_indent_overrides_auto(self):
        text = console.render_table(("A",), [("1",)], row_prefixes=["X "], indent="")
        assert text.splitlines()[0].startswith("A")

    def test_column_count_mismatch_raises(self):
        with pytest.raises(ValueError, match="列数不一致"):
            console.render_table(("A", "B"), [("1",)])

    def test_max_widths_length_mismatch_raises(self):
        with pytest.raises(ValueError, match="max_widths"):
            console.render_table(("A", "B"), [("1", "2")], max_widths=(3,))

    def test_row_prefixes_length_mismatch_raises(self):
        with pytest.raises(ValueError, match="row_prefixes"):
            console.render_table(("A",), [("1",), ("2",)], row_prefixes=["x"])

    def test_empty_rows(self):
        text = console.render_table(("A", "B"), [])
        assert len(text.splitlines()) == 2


class TestStreamSafety:
    def test_ensure_safe_output_does_not_raise(self):
        console.ensure_safe_output()

    def test_console_encoding_returns_string(self):
        assert isinstance(console.console_encoding(), str)
        assert console.console_encoding()

    def test_does_not_switch_encoding_to_utf8(self, monkeypatch):
        """回归测试：绝不能把 stdout 改成 UTF-8，否则中文 Windows 控制台会乱码。"""
        calls: list[dict] = []

        class FakeStream:
            encoding = "cp936"

            def reconfigure(self, **kwargs):
                calls.append(kwargs)

        monkeypatch.setattr("sys.stdout", FakeStream())
        monkeypatch.setattr("sys.stderr", FakeStream())
        console.ensure_safe_output()

        assert calls, "应当调用 reconfigure 设置 errors"
        for kwargs in calls:
            assert "encoding" not in kwargs
            assert kwargs.get("errors") == "replace"


class TestPrintHelpers:
    def test_heading_prints_banner(self, capsys):
        console.heading("标题", width=10)
        out = capsys.readouterr().out.splitlines()
        assert out[0] == "=" * 10
        assert out[1] == "标题"
        assert out[2] == "=" * 10

    def test_status_helpers_have_prefix(self, capsys):
        console.ok("好")
        console.warn("注意")
        console.fail("坏了")
        console.info("提示")
        out = capsys.readouterr().out
        assert "[OK]" in out
        assert "[警告]" in out
        assert "[失败]" in out
        assert "[信息]" in out
