"""工具：生成采购清单（STEP 模型 + PDF 图纸 → 采购清单 Excel）。

流程
----
1. 选目标文件夹（命令行给 / 弹窗选 / 手输）
2. 扫出文件夹（含子目录）里所有 ``.stp`` / ``.step``
3. 为每个模型找**同目录同名** PDF，从图纸里读两样东西：
   * 标题栏的**材料**
   * 「技术要求 / 设计要求」栏里的**表面处理要求**（写进备注列）
4. 生成 ``<文件夹名>.xlsx``，写名称、材料、数量、备注
5. 汇总报告：哪些没图纸、哪些有图纸但没材料、哪些图纸没写表面处理

对应原脚本：``D:\\WorkSpace\\订单\\生成采购清单.py``（v1.0）。
"""

from __future__ import annotations

import argparse
import os
import tempfile
from functools import partial
from pathlib import Path

from .. import console
from ..config import MODEL_EXTS
from ..dialogs import ask_path, pick_folder, validate_folder
from ..errors import InputError, SpeedupError
from ..excel.writer import generate_supplement_list
from ..material import index_pdf_by_stem, lookup_material_for
from ..models import SupplementEntry
from ..osutil import open_with_default_app
from ..techreq import find_craft_in_pdf
from ..thumb import is_available, render_isometric

NAME = "gbbuild"
TITLE = "生成采购清单"
SUMMARY = "扫文件夹里的 STP 模型，读同名 PDF 图纸的材料和表面处理要求，生成采购清单 Excel"


def add_parser(subparsers: argparse._SubParsersAction) -> argparse.ArgumentParser:
    parser = subparsers.add_parser(
        NAME,
        help=SUMMARY,
        description=f"{TITLE}：{SUMMARY}。不带参数运行会弹窗让你选文件夹。",
        epilog="示例：speedup gbbuild \"D:\\WorkSpace\\订单\\0916\" --no-open",
    )
    parser.add_argument(
        "folder",
        nargs="?",
        default=None,
        help="目标文件夹；省略则弹窗选择",
    )
    parser.add_argument("--out", default=None, help="输出文件名（默认用文件夹名 + .xlsx）")
    parser.add_argument("--no-open", action="store_true", help="生成后不要自动打开表格")
    parser.add_argument(
        "--with-image",
        action="store_true",
        help="给每个零件渲染轴测图填进「图片」列（需要可选的 cadquery 依赖，慢）",
    )
    return parser


def _choose_folder(initial_dir: str | None) -> str:
    """命令行没给文件夹时：弹窗 → 手输 → 取消。"""
    chosen = ask_path(
        "folder",
        "请输入目标文件夹路径：",
        partial(pick_folder, "选择要生成采购清单的文件夹"),
        validate_folder,
        initial_dir=initial_dir,
    )
    if not chosen:
        raise SpeedupError("用户未选择文件夹")
    return chosen


def collect_entries(
    folder: str, *, with_image: bool = False
) -> tuple[list[SupplementEntry], dict[str, list[str]]]:
    """扫描模型 + 识别材料 + 读图纸技术要求栏 + （可选）渲染轴测图。

    :param with_image: 是否给每个零件渲染轴测图（需要可选的 cadquery 依赖，
        每个零件约 5 秒）
    :returns: ``(零件列表, {"missing_pdf": [...], "no_material": [...],
        "no_craft": [...], "lookup_failed": [...], "image_failed": [...]})``
    """
    models: list[tuple[str, str]] = []
    for dirpath, _dirnames, filenames in os.walk(folder):
        for filename in filenames:
            if filename.startswith(("~$", ".")):
                continue
            stem, ext = os.path.splitext(filename)
            if ext.lower() in MODEL_EXTS:
                models.append((os.path.join(dirpath, filename), stem))
    models.sort(key=lambda item: item[1].lower())

    if not models:
        raise InputError(f"文件夹（含子目录）里没有找到任何 {' / '.join(MODEL_EXTS)} 文件：{folder}")

    pdf_index = index_pdf_by_stem(folder)
    # 轴测图丢在系统临时目录：反正会被嵌进 xlsx，不必留在用户文件夹里添乱（R23）
    image_dir = (
        Path(tempfile.mkdtemp(prefix="speedup-thumbs-")) if with_image else None
    )
    if with_image and not is_available():
        print("[警告] 没装 cadquery，跳过轴测图（pip install speedup[images] 可补上）")

    entries: list[SupplementEntry] = []
    missing_pdf: list[str] = []
    no_material: list[str] = []
    no_craft: list[str] = []
    lookup_failed: list[str] = []
    image_failed: list[str] = []

    for path, stem in models:
        lookup, pdf_path = lookup_material_for(path, pdf_index)
        if pdf_path is None:
            missing_pdf.append(stem)
        elif lookup.error:
            lookup_failed.append(f"{stem}({lookup.error})")
        elif lookup.material is None:
            no_material.append(stem)

        # 表面处理要求（写进备注列）。
        #
        # 图纸本身就读不了的时候**不再读第二次** —— 同一份 PDF 失败两遍，
        # 报告里会出现两条重复的警告，反而看不出到底哪里坏了。
        note: str | None = None
        if pdf_path is not None and not lookup.error:
            craft = find_craft_in_pdf(pdf_path)
            note = craft.text
            if note is None:
                if craft.error:
                    # 「读不了」和「图纸没写」是两回事，分开报（见 AGENTS.md R26）
                    lookup_failed.append(f"{stem}(表面处理:{craft.error})")
                else:
                    no_craft.append(stem)

        image: str | None = None
        if image_dir is not None:
            result = render_isometric(path, image_dir / f"{stem}.png")
            if result.ok:
                image = result.path
            else:
                image_failed.append(f"{stem}({result.error})")

        entries.append(
            SupplementEntry(
                name=stem, material=lookup.material, note=note, image=image
            )
        )

    return entries, {
        "missing_pdf": missing_pdf,
        "no_material": no_material,
        "no_craft": no_craft,
        "lookup_failed": lookup_failed,
        "image_failed": image_failed,
    }


def run(args: argparse.Namespace) -> int:
    console.ensure_safe_output()

    folder = args.folder or _choose_folder(os.path.dirname(os.path.abspath(__file__)))
    folder = os.path.abspath(folder)
    if not os.path.isdir(folder):
        raise InputError(f"文件夹不存在：{folder}")

    console.heading(f"{TITLE}  ——  {folder}")

    print("\n正在扫描模型文件...")
    if args.with_image:
        print("（已开启轴测图渲染：每个零件约 5 秒，请稍候）")
    entries, notes = collect_entries(folder, with_image=args.with_image)
    print(f"共找到 {len(entries)} 个模型文件。")

    print("\n识别结果：")
    console.print_table(
        ("序号", "名称", "材料", "表面处理", "材料来源"),
        [
            (
                str(index),
                entry.name,
                entry.material or "-",
                entry.note or "-",
                _source_of(entry, notes),
            )
            for index, entry in enumerate(entries, 1)
        ],
        max_widths=(4, 24, 11, 20, 8),
    )

    destination, existed = generate_supplement_list(folder, entries, args.out)

    console.section("生成结果")
    console.ok(f"已生成：{destination}" + ("（覆盖旧文件）" if existed else ""))
    print(f"共写入 {len(entries)} 个零件。")
    _print_notes(notes, folder)

    if not args.no_open:
        open_with_default_app(destination)
    return 0


def _source_of(entry: SupplementEntry, notes: dict[str, list[str]]) -> str:
    if entry.material:
        return "已识别"
    if entry.name in notes["missing_pdf"]:
        return "无同名图纸"
    if entry.name in notes["no_material"]:
        return "图纸内未找到"
    return "读取失败"


def _print_notes(notes: dict[str, list[str]], folder: str) -> None:
    missing = notes["missing_pdf"]
    no_material = notes["no_material"]
    no_craft = notes["no_craft"]
    failed = notes["lookup_failed"]
    image_failed = notes["image_failed"]

    if missing:
        console.warn(f"{len(missing)} 个零件没有同名图纸（材料、备注留空）：{'、'.join(missing)}")
        _explain_missing_pdfs(folder)
    if no_material:
        console.warn(
            f"{len(no_material)} 个零件有图纸但没识别出材料（材料留空）：{'、'.join(no_material)}"
        )
    if no_craft:
        console.warn(
            f"{len(no_craft)} 个零件的图纸技术要求栏里没有写表面处理（备注留空）："
            f"{'、'.join(no_craft)}"
        )
    if failed:
        console.warn(f"{len(failed)} 个零件的图纸读取失败：{'、'.join(failed)}")
    if image_failed:
        console.warn(
            f"{len(image_failed)} 个零件的轴测图渲染失败（图片留空）：{'、'.join(image_failed)}"
        )
    if not (missing or no_material or no_craft or failed or image_failed):
        console.ok("所有零件的材料、表面处理和轴测图都已就绪。")


def _explain_missing_pdfs(folder: str) -> None:
    """列出文件夹里实际有哪些 PDF。

    匹配规则是「**同一个文件夹**里同名」—— 用户手上有一张图纸、却被报成
    「没有同名图纸」时，八成是两种情况之一：图纸根本不在这个文件夹，
    或者图纸在、但文件名和模型对不上。把文件夹里的 PDF 列出来，一眼就能分清。
    """
    names: list[str] = []
    for _dirpath, _dirnames, filenames in os.walk(folder):
        for filename in filenames:
            if os.path.splitext(filename)[1].lower() == ".pdf":
                names.append(filename)

    if not names:
        print("        这个文件夹（含子目录）里一张 PDF 都没有 —— 图纸是不是放在别处了？")
        return
    shown = "、".join(sorted(names))
    print(f"        这个文件夹（含子目录）里共有 {len(names)} 张 PDF：{shown}")
    print("        上面缺的零件若不在这批文件名里，就是图纸和模型的名字对不上。")
