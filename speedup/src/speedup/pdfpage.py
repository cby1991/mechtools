"""pdfplumber 页面读取的公共部分。

为什么要有这一层
----------------
**只取落在页面范围内的词。** PDF 的页面之外可能还有内容 —— 两张图拼在一起、
导出时画布比页面大，等等。pdfplumber 会照样把它们 ``extract_words()`` 出来，
而它们的坐标（``x0`` 超出页宽）在屏幕上根本看不见。

实测踩到：一张 595pt 宽的图纸上，``x0=627.9`` 处藏着**另一套「技术要求」栏**
（写着「外观件，表面电镀处理」），而页内那套写的是「表面本色硬质阳极氧化」。
因为两套词在 ``extract_words()`` 里的先后顺序不稳定，同一个目录下 5 张图纸里
有 3 张恰好读对、2 张读成了页面外那套 —— 而且**结果看起来完全合理**
（「电镀」也是个合法的表面处理），不回去核对原图根本发现不了。

所以读图纸一律走这里，别直接 ``page.extract_words()``。
"""

from __future__ import annotations

from typing import Any

__all__ = ["words_in_page"]

#: 边缘容差。PDF 里的坐标常带零点几 pt 的浮点误差，卡太死会误杀贴着页边的字。
_EDGE_TOLERANCE = 0.5


def words_in_page(page: Any) -> list[Any]:
    """取出**落在页面范围内**的词。

    页面外的内容在图纸上是看不见的（被裁剪掉了），但它们仍留在 PDF 的内容流里。
    当成有效内容读进来，就会得到「图上没有、却又合情合理」的结果。

    :param page: pdfplumber 的 ``Page`` 对象
    :returns: 词列表，字段与 ``page.extract_words()`` 一致
    """
    left, top = 0.0, 0.0
    right, bottom = float(page.width), float(page.height)
    return [
        word
        for word in page.extract_words()
        if word["x0"] >= left - _EDGE_TOLERANCE
        and word["x1"] <= right + _EDGE_TOLERANCE
        and word["top"] >= top - _EDGE_TOLERANCE
        and word["bottom"] <= bottom + _EDGE_TOLERANCE
    ]
