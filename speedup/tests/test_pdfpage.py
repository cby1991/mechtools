"""``words_in_page`` 的测试：页面外的词必须被丢掉。"""

from __future__ import annotations

from speedup.pdfpage import words_in_page


class FakePage:
    """够用的 stub：只需要 ``width`` / ``height`` / ``extract_words()``。"""

    def __init__(self, words, width=595.0, height=842.0):
        self._words = words
        self.width = width
        self.height = height

    def extract_words(self):
        return list(self._words)


def word(text, x0, x1, top, bottom):
    return {"text": text, "x0": x0, "x1": x1, "top": top, "bottom": bottom}


class TestWordsInPage:
    def test_drops_words_beyond_the_right_edge(self):
        """回归：595pt 宽的图纸上，x0=627.9 处还藏着另一套「技术要求」。

        那是被裁掉的部分，屏幕上看不见，绝不能当成有效内容 ——
        实测它会把「表面本色硬质阳极氧化」顶替成「外观件，表面电镀处理」。
        """
        page = FakePage([
            word("技术要求", 79.4, 130.0, 553.9, 566.0),
            word("技术要求：", 627.9, 680.0, 523.8, 536.0),
        ])
        kept = words_in_page(page)
        assert [w["text"] for w in kept] == ["技术要求"]

    def test_drops_words_beyond_the_bottom_edge(self):
        page = FakePage([word("低", 10.0, 40.0, 850.0, 870.0)])
        assert words_in_page(page) == []

    def test_drops_words_off_the_left_edge(self):
        page = FakePage([word("左", -30.0, -5.0, 100.0, 120.0)])
        assert words_in_page(page) == []

    def test_keeps_words_touching_the_edges(self):
        """贴着页边的字不能被误杀（PDF 里常有零点几 pt 的浮点误差）。"""
        page = FakePage([word("满", 0.0, 595.0, 0.0, 842.0)])
        assert len(words_in_page(page)) == 1

    def test_keeps_normal_words(self):
        page = FakePage([word("正常", 100.0, 140.0, 200.0, 220.0)])
        assert len(words_in_page(page)) == 1
