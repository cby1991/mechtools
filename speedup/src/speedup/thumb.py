"""给零件生成轴测图（采购清单「图片」列用）。

为什么单独一个模块
------------------
渲染依赖 **cadquery / OCP**（几百 MB，还捎带 vtk / numpy / scipy / numba），
所以它是**可选**的：没装时 :func:`is_available` 返回 ``False``、
:func:`render_isometric` 返回带 ``error`` 的结果，其它功能完全不受影响。

分层（照 `material.py` 的约定）
--------------------------------
* :func:`isometric_camera` / :func:`fit_to_cell` —— 纯函数，不碰磁盘、不装依赖
* :func:`render_isometric` —— 唯一需要 cadquery 的一层

两个必须这么做的理由
--------------------
**1. 相机必须自己算，不能用 ``azimuth/elevation``。**
实测 ``show(azimuth=45, elevation=35.264, viewup=(0,0,1))`` 会让 VTK 报
``Resetting view-up since view plane normal is parallel`` 并把视角重置成俯视；
``orthographic=True`` 更是直接出空白图（试了 4 组参数全是 1725 字节）。
**只有 ``position=`` 手动定相机是可靠的**，所以位置在这里算好。

**2. 零件必须裁到撑满画面。**
表格里「图片」列只有 126px 宽，留白多一点点零件就小得看不清。
但"用固定相机距离"不可靠 —— 实测同一个距离系数下，圆润零件填充率 96%、
长条件只有 74%。所以：**渲染一个大画布，再按非白区域裁紧、补一圈很小的边距**，
这样不论什么形状，零件都稳定占满画面。
"""

from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

__all__ = [
    "ThumbResult",
    "GRAY_COLOR",
    "is_available",
    "isometric_camera",
    "fit_to_cell",
    "render_isometric",
]

#: 零件本体颜色。STEP AP203 不带颜色信息，得我们自己定。
GRAY_COLOR = "gray"

#: 渲染画布尺寸。比目标大得多 —— 反正最后要缩到 126px 宽，大画布换来裁切余量。
CANVAS = 800

#: 相机距离 = 包围盒对角线 × 这个系数。
#: 1.6 是留足余量不切角；真正的"撑满"交给裁白边去做，不靠调这个数。
CAMERA_FACTOR = 1.6

#: 裁紧之后补的边距（占短边的比例）。3% 刚好不贴死，又不浪费像素。
MARGIN_RATIO = 0.03

#: 判定"这是背景"的阈值：三通道之和大于它就算白。
_WHITE_SUM = 700


@dataclass(frozen=True, slots=True)
class ThumbResult:
    """一次渲染的结果。``error`` 非空表示没成功。"""

    path: str | None = None
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.path is not None


def _load_cadquery() -> Any | None:
    """拿到 cadquery 模块；没装返回 ``None``（不抛异常）。"""
    try:
        import cadquery  # noqa: PLC0415  —— 重依赖，必须延迟导入
    except ImportError:
        return None
    return cadquery


def is_available() -> bool:
    """渲染能力是否可用（装了 cadquery 才可用）。"""
    return _load_cadquery() is not None


def isometric_camera(
    diagonal: float, factor: float = CAMERA_FACTOR
) -> tuple[float, float, float]:
    """算等轴测的相机位置。

    视线方向是 ``(+X, -Y, +Z)`` —— 也就是 Z 朝上、X 往左下、Y 往右下，
    与 SolidWorks 的等轴测一致。三个分量绝对值相等，才是标准等轴测
    （三根轴在画面里各夹 120°）。

    :param diagonal: 零件包围盒的空间对角线长度
    :param factor: 距离系数，越大零件越小（1.6 留足余量不切角）
    """
    k = 1 / (3**0.5)
    distance = diagonal * factor
    return (distance * k, -distance * k, distance * k)


def fit_to_cell(
    image_size: tuple[int, int], cell_size: tuple[int, int]
) -> tuple[int, int]:
    """算图片塞进单元格后的显示尺寸（保持比例、最大内接、不溢出）。

    :param image_size: 原图 ``(宽, 高)``
    :param cell_size: 单元格 ``(宽, 高)``（像素）
    """
    image_w, image_h = image_size
    cell_w, cell_h = cell_size
    if image_w <= 0 or image_h <= 0:
        return (1, 1)
    scale = min(cell_w / image_w, cell_h / image_h)
    return (max(1, int(image_w * scale)), max(1, int(image_h * scale)))


def _crop_to_content(image: Any, margin_ratio: float = MARGIN_RATIO) -> Any:
    """按非白区域裁紧，再补一圈很小的边距。

    这是"让零件撑满画面"的关键一步：相机距离只能保证零件**完整**，
    真正把它放大到贴边靠的是这里。
    """
    from PIL import Image, ImageChops

    background = Image.new("RGB", image.size, (255, 255, 255))
    box = ImageChops.difference(image, background).getbbox()
    if box is None:  # 整张都是白的 —— 零件没渲染出来
        return image

    cropped = image.crop(box)
    pad = max(2, int(min(cropped.size) * margin_ratio))
    padded = Image.new(
        "RGB", (cropped.size[0] + pad * 2, cropped.size[1] + pad * 2), (255, 255, 255)
    )
    padded.paste(cropped, (pad, pad))
    return padded


def render_isometric(  # noqa: PLR0913
    step_path: str | os.PathLike[str],
    out_path: str | os.PathLike[str],
    *,
    canvas: int = CANVAS,
    color: str = GRAY_COLOR,
    factor: float = CAMERA_FACTOR,
    margin_ratio: float = MARGIN_RATIO,
) -> ThumbResult:
    """把 STEP 渲染成等轴测 PNG。

    读不了、渲染失败时返回带 ``error`` 的结果，**不抛异常** —— 与
    :func:`speedup.material.find_material_in_pdf` 的约定一致，调用方只看
    :attr:`ThumbResult.ok`。

    :param step_path: ``.stp`` / ``.step`` 文件
    :param out_path: 输出的 PNG；父目录不存在会自动创建
    """
    cq = _load_cadquery()
    if cq is None:
        return ThumbResult(error="未安装 cadquery（可选依赖）")

    step_path, out_path = Path(step_path), Path(out_path)
    if not step_path.exists():
        return ThumbResult(error=f"文件不存在：{step_path}")

    try:
        from cadquery.vis import show, style
    except ImportError as exc:  # cadquery 在、但 vis 缺（缺 vtk）
        return ThumbResult(error=f"缺少可视化组件({exc})")

    try:
        shape = cq.importers.importStep(str(step_path))
        bound = shape.val().BoundingBox()
        # 挪到原点：相机是按原点算的，零件留在原位会跑到画面外（实测出过空图）
        shape = shape.translate((-bound.center.x, -bound.center.y, -bound.center.z))
        diagonal = (
            bound.xlen**2 + bound.ylen**2 + bound.zlen**2
        ) ** 0.5
    except Exception as exc:
        return ThumbResult(error=f"读取 STEP 失败({exc})")

    out_path.parent.mkdir(parents=True, exist_ok=True)

    # 临时文件写到系统临时目录，不落在用户目录里（R23）
    with tempfile.TemporaryDirectory(prefix="speedup-thumb-") as tmp:
        raw = Path(tmp) / "raw.png"
        try:
            show(
                style(shape, color=color),
                width=canvas,
                height=canvas,
                screenshot=str(raw),
                position=isometric_camera(diagonal, factor),
                interact=False,
                trihedron=False,  # 表格里不需要坐标轴三角
                edges=False,  # 打开会画出三角剖分网格，很难看
                gradient=False,  # 纯白底，便于裁切和打印
            )
        except Exception as exc:
            return ThumbResult(error=f"渲染失败({exc})")

        if not raw.exists():
            return ThumbResult(error="渲染没有产生图像")

        try:
            from PIL import Image

            image = Image.open(raw).convert("RGB")
            _crop_to_content(image, margin_ratio).save(out_path)
        except Exception as exc:
            return ThumbResult(error=f"裁剪/保存失败({exc})")

    return ThumbResult(path=str(out_path))
