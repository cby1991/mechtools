"""跨模块共享的领域数据结构。"""

from __future__ import annotations

from dataclasses import dataclass

from .config import DEFAULT_QTY

__all__ = ["SupplementEntry"]


@dataclass(frozen=True, slots=True)
class SupplementEntry:
    """采购清单里的一行零件。

    :param name: 零件名称（写进「名称」列）
    :param material: 材料牌号；识别不到时留空
    :param qty: 数量，默认 1
    :param note: 备注。放的是从图纸「技术要求 / 设计要求」栏读到的**表面处理要求**
        （如「喷砂本色阳极氧化」）；图纸没写、或者根本没有同名图纸时留空。
        备注列只有一列 —— 以后若还要塞别的信息进去，**必须在这里拼起来**，别覆盖。
    :param image: 轴测图 PNG 路径（写进「图片」列）；没开渲染或渲染失败时留空
    """

    name: str
    material: str | None = None
    qty: int = DEFAULT_QTY
    note: str | None = None
    image: str | None = None
