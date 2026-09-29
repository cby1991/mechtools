"""生成 speedup 的图标 ``src/speedup/assets/speedup.ico``。

图标要给 Windows 快捷方式用（``shortcut.py`` 创建 .lnk 时会引用它），
用 Pillow 画一次就好。改配色或图形后重新跑一遍：

    uv run python examples/make_icon.py

和 ``make_placeholder_images.py`` 一样，这个脚本**不属于任何工具的运行路径**，
只是素材生成器；跑一次产出物就提交进仓库，日常不会再执行。
"""

from __future__ import annotations

import os

from PIL import Image, ImageDraw

#: 和 gui.py / tools/poster.py 保持同一套配色
NAVY = (20, 54, 92)
GOLD = (226, 163, 61)

#: 闪电形状，坐标是 0~1 的相对值，缩放后画进圆角方块里
BOLT = (
    (0.56, 0.07),
    (0.22, 0.57),
    (0.45, 0.57),
    (0.36, 0.93),
    (0.74, 0.41),
    (0.50, 0.41),
)

#: Windows 会按需要挑最合适的一档，16 和 32 是任务栏/资源管理器最常用的
SIZES = (16, 24, 32, 48, 64, 128, 256)

#: 画布边长，其他尺寸由 Pillow 降采样得到
CANVAS = 256

HERE = os.path.dirname(os.path.abspath(__file__))
OUTPUT = os.path.join(os.path.dirname(HERE), "src", "speedup", "assets", "speedup.ico")


def build() -> Image.Image:
    image = Image.new("RGBA", (CANVAS, CANVAS), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    # 圆角方块：留一点边，避免图标在任务栏里显得过满
    margin = CANVAS * 0.06
    draw.rounded_rectangle(
        [margin, margin, CANVAS - margin, CANVAS - margin],
        radius=CANVAS * 0.22,
        fill=NAVY,
    )

    draw.polygon([(x * CANVAS, y * CANVAS) for x, y in BOLT], fill=GOLD)
    return image


def main() -> int:
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    image = build()
    image.save(OUTPUT, format="ICO", sizes=[(size, size) for size in SIZES])
    size_kb = os.path.getsize(OUTPUT) / 1024
    print(f"已生成：{OUTPUT}")
    print(f"包含尺寸：{', '.join(f'{s}×{s}' for s in SIZES)}，共 {size_kb:.1f} KB")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
