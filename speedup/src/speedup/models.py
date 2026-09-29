"""跨模块共享的领域数据结构。"""

from __future__ import annotations

from dataclasses import dataclass

from .config import DEFAULT_QTY

__all__ = ["SupplementEntry"]


@dataclass(frozen=True, slots=True)
class SupplementEntry:
    """补料清单里的一行零件。

    :param name: 零件名称（写进「名称」列）
    :param material: 材料牌号；识别不到时留空
    :param qty: 数量，默认 1
    :param note: 备注，默认空（例如「未找到同名图纸」）
    """

    name: str
    material: str | None = None
    qty: int = DEFAULT_QTY
    note: str | None = None
