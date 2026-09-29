"""工具：生成补料清单（STEP 模型 + PDF 图纸 → 补料清单 Excel）。

流程
----
1. 选目标文件夹（命令行给 / 弹窗选 / 手输）
2. 扫出文件夹（含子目录）里所有 ``.stp`` / ``.step``
3. 为每个模型找**同目录同名** PDF，从图纸标题栏识别材料
4. 生成 ``<文件夹名>.xlsx``，写名称、材料、数量
5. 汇总报告：哪些没图纸、哪些有图纸但没识别出材料

对应原脚本：``D:\\WorkSpace\\订单\\生成补料清单.py``（v1.0）。
"""

from __future__ import annotations

import argparse
import os
from functools import partial

from .. import console
from ..config import MODEL_EXTS
from ..dialogs import ask_path, pick_folder, validate_folder
from ..errors import InputError, SpeedupError
from ..excel.writer import generate_supplement_list
from ..material import index_pdf_by_stem, lookup_material_for
from ..models import SupplementEntry
from ..osutil import open_with_default_app

NAME = "gbbuild"
TITLE = "生成补料清单"
SUMMARY = "扫文件夹里的 STP 模型，读同名 PDF 图纸标题栏的材料，生成补料清单 Excel"


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
    return parser


def _choose_folder(initial_dir: str | None) -> str:
    """命令行没给文件夹时：弹窗 → 手输 → 取消。"""
    chosen = ask_path(
        "folder",
        "请输入目标文件夹路径：",
        partial(pick_folder, "选择要生成补料清单的文件夹"),
        validate_folder,
        initial_dir=initial_dir,
    )
    if not chosen:
        raise SpeedupError("用户未选择文件夹")
    return chosen


def collect_entries(folder: str) -> tuple[list[SupplementEntry], dict[str, list[str]]]:
    """扫描模型 + 识别材料。

    :returns: ``(零件列表, {"missing_pdf": [...], "no_material": [...]})``
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

    entries: list[SupplementEntry] = []
    missing_pdf: list[str] = []
    no_material: list[str] = []
    lookup_failed: list[str] = []

    for path, stem in models:
        lookup, pdf_path = lookup_material_for(path, pdf_index)
        if pdf_path is None:
            missing_pdf.append(stem)
        elif lookup.error:
            lookup_failed.append(f"{stem}({lookup.error})")
        elif lookup.material is None:
            no_material.append(stem)
        entries.append(SupplementEntry(name=stem, material=lookup.material))

    return entries, {
        "missing_pdf": missing_pdf,
        "no_material": no_material,
        "lookup_failed": lookup_failed,
    }


def run(args: argparse.Namespace) -> int:
    console.ensure_safe_output()

    folder = args.folder or _choose_folder(os.path.dirname(os.path.abspath(__file__)))
    folder = os.path.abspath(folder)
    if not os.path.isdir(folder):
        raise InputError(f"文件夹不存在：{folder}")

    console.heading(f"{TITLE}  ——  {folder}")

    print("\n正在扫描模型文件...")
    entries, notes = collect_entries(folder)
    print(f"共找到 {len(entries)} 个模型文件。")

    print("\n材料识别结果：")
    console.print_table(
        ("序号", "名称", "材料", "来源"),
        [
            (str(index), entry.name, entry.material or "-", _source_of(entry, notes))
            for index, entry in enumerate(entries, 1)
        ],
        max_widths=(4, 40, 14, 12),
    )

    destination, existed = generate_supplement_list(folder, entries, args.out)

    console.section("生成结果")
    console.ok(f"已生成：{destination}" + ("（覆盖旧文件）" if existed else ""))
    print(f"共写入 {len(entries)} 个零件。")
    _print_notes(notes)

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


def _print_notes(notes: dict[str, list[str]]) -> None:
    missing = notes["missing_pdf"]
    no_material = notes["no_material"]
    failed = notes["lookup_failed"]

    if missing:
        console.warn(f"{len(missing)} 个零件没有同名图纸（材料留空）：{'、'.join(missing)}")
    if no_material:
        console.warn(
            f"{len(no_material)} 个零件有图纸但没识别出材料（材料留空）：{'、'.join(no_material)}"
        )
    if failed:
        console.warn(f"{len(failed)} 个零件的图纸读取失败：{'、'.join(failed)}")
    if not (missing or no_material or failed):
        console.ok("所有零件的材料都已识别。")
