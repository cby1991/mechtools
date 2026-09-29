"""轴测图渲染测试。

分两层：
* **纯函数**（相机位置、嵌入尺寸）—— 不装 cadquery 也能跑
* **渲染**（真出 PNG）—— 用 ``importorskip`` 保护，没装就跳过（R16 精神）
"""

from __future__ import annotations

import pytest
from PIL import Image

from speedup.thumb import (
    ThumbResult,
    fit_to_cell,
    is_available,
    isometric_camera,
    render_isometric,
)

CELL = (126, 95)  # 「图片」列按 18 字符宽、100pt 行高折算出来的像素


class TestIsometricCamera:
    def test_direction_is_plus_x_minus_y_plus_z(self):
        """等轴测朝向：X 正 / Y 负 / Z 正。

        这是主人在 SolidWorks 里截图确认过的方向 —— Z 朝上、X 往左下、Y 往右下。
        改这个符号等于把视角翻到侧面去。
        """
        x, y, z = isometric_camera(100.0)
        assert x > 0
        assert y < 0
        assert z > 0

    def test_three_axes_are_equal_length(self):
        """三个分量绝对值相等 = 标准等轴测（三轴投影各 120°）。

        只要不等，画出来就不是等轴测而是某个斜视图。
        """
        x, y, z = isometric_camera(100.0)
        assert abs(abs(x) - abs(y)) < 1e-9
        assert abs(abs(y) - abs(z)) < 1e-9

    def test_distance_grows_with_diagonal(self):
        """零件越大、相机越远 —— 这样不同尺寸的零件在画面里大小才一致。"""
        near = isometric_camera(10.0)
        far = isometric_camera(1000.0)
        assert sum(c * c for c in far) > sum(c * c for c in near)

    def test_factor_zooms_linearly(self):
        a = isometric_camera(100.0, factor=1.0)
        b = isometric_camera(100.0, factor=2.0)
        da = sum(c * c for c in a) ** 0.5
        db = sum(c * c for c in b) ** 0.5
        assert abs(db / da - 2.0) < 1e-9

    def test_accepts_zero_diagonal(self):
        """退化输入不该抛异常。"""
        assert isometric_camera(0.0) == (0.0, -0.0, 0.0)


class TestFitToCell:
    """把渲染出来的图塞进单元格时，算显示尺寸（保持比例、最大内接）。"""

    def test_wide_image_limited_by_width(self):
        assert fit_to_cell((400, 100), CELL) == (126, 31)

    def test_tall_image_limited_by_height(self):
        assert fit_to_cell((100, 400), CELL) == (23, 95)

    def test_square_limited_by_the_shorter_side(self):
        assert fit_to_cell((200, 200), CELL) == (95, 95)

    @pytest.mark.parametrize(
        "size", [(1, 1), (10, 500), (500, 10), (333, 222), (126, 95)]
    )
    def test_never_overflows_the_cell(self, size):
        width, height = fit_to_cell(size, CELL)
        assert width <= CELL[0]
        assert height <= CELL[1]
        assert width >= 1 and height >= 1

    @pytest.mark.parametrize("size", [(400, 300), (714, 584), (100, 400)])
    def test_keeps_aspect_ratio(self, size):
        width, height = fit_to_cell(size, CELL)
        assert abs(width / height - size[0] / size[1]) < 0.05


class TestIsAvailable:
    def test_returns_bool(self):
        assert isinstance(is_available(), bool)

    def test_false_when_cadquery_missing(self, monkeypatch):
        """没装 cadquery 时必须是 False，而不是抛异常。"""
        monkeypatch.setattr("speedup.thumb._load_cadquery", lambda: None)
        assert is_available() is False


class TestRenderWithoutCadquery:
    def test_reports_error_instead_of_crashing(self, tmp_path, monkeypatch):
        """缺依赖要给 error，不能崩 —— 调用方只看得懂 ThumbResult。"""
        monkeypatch.setattr("speedup.thumb._load_cadquery", lambda: None)
        result = render_isometric(tmp_path / "不存在.step", tmp_path / "out.png")
        assert isinstance(result, ThumbResult)
        assert result.ok is False
        assert result.error


class TestRenderIsometric:
    """真渲染。没装 cadquery 的环境自动跳过。"""

    @pytest.fixture
    def sample_step(self, tmp_path):
        """现造一个零件再导出 STEP —— 不碰任何真实图纸（R13）。"""
        cq = pytest.importorskip("cadquery")
        box = cq.Workplane().box(20, 12, 8)
        path = tmp_path / "试件.step"
        cq.exporters.export(box, str(path))
        return path

    def test_produces_a_png(self, sample_step, tmp_path):
        out = tmp_path / "thumb.png"
        result = render_isometric(sample_step, out)
        assert result.ok, result.error
        assert out.exists()
        assert Image.open(out).size[0] > 0

    def test_whitespace_is_cropped_away(self, sample_step, tmp_path):
        """零件要撑满画面 —— 四周不能留着大片白边。

        做法是渲染大画布再按非白区域裁紧；留着白边的话，
        表格里那个 126px 宽的小格里零件会小得看不清。
        """
        out = tmp_path / "thumb.png"
        render_isometric(sample_step, out)
        image = Image.open(out).convert("RGB")
        # 四角必须已经是白的（裁紧之后才可能这样），中心必须有内容
        corners = [
            image.getpixel((0, 0)),
            image.getpixel((image.width - 1, 0)),
            image.getpixel((0, image.height - 1)),
            image.getpixel((image.width - 1, image.height - 1)),
        ]
        assert all(sum(c) > 700 for c in corners), "四角应为白底"
        assert image.getpixel((image.width // 2, image.height // 2)) != (255, 255, 255)

    def test_missing_step_reports_error(self, tmp_path):
        pytest.importorskip("cadquery")
        result = render_isometric(tmp_path / "没有这个.step", tmp_path / "o.png")
        assert result.ok is False
        assert result.error

    def test_creates_parent_directory(self, sample_step, tmp_path):
        """输出的父目录不存在时要自己建，否则调用方得先 mkdir。"""
        out = tmp_path / "子目录" / "深一层" / "thumb.png"
        result = render_isometric(sample_step, out)
        assert result.ok, result.error
        assert out.exists()
