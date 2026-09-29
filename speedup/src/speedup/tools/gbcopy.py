"""工具：补料文件自动匹配拷贝。

流程
----
1. 选补料清单 Excel（命令行给 / 弹窗选 / 手输）
2. 读出「名称」列的所有零件名（**只读一张工作表**，见下）
3. 选在哪个文件夹里搜索
4. 扫描其中的 ``.stp/.step/.pdf``，按名称匹配
5. 列出匹配结果供核对，然后拷到清单所在文件夹（**不重命名、不覆盖同名文件**）

只读一张工作表
--------------
工作簿里常常还有「备料参考」这类同样带「名称」列的附页。
默认**优先读名为「补料清单」的那张**，没有就取第一张带「名称」列的表。
有别的候选表时会打一条醒目提醒，用 ``--sheet`` 可以手动指定。

> 早期实现会把所有表的名合并，导致参考页的零件也被拷出去。见 `AGENTS.md` 的 P12。

特点
----
* 支持文件名带版本号：清单写「驱动连杆」，文件叫「驱动连杆V1.1.STEP」也能匹配
* 扩展名、大小写不敏感
* ``--dry-run`` 只列结果不拷贝

对应原脚本：``D:\\WorkSpace\\订单\\补料文件自动匹配拷贝.py``（v1.0）。
"""

from __future__ import annotations

import argparse
import os
from collections.abc import Sequence
from functools import partial

from .. import console
from ..config import DEFAULT_MIN_SCORE
from ..dialogs import (
    FILTER_EXCEL,
    ask_path,
    pick_file,
    pick_folder,
    validate_file,
    validate_folder,
)
from ..errors import SpeedupError
from ..excel.reader import extract_names
from ..matching import (
    LEVEL_MARKERS,
    CopyAction,
    CopyReport,
    MatchResult,
    build_matches,
    execute_copies,
    find_shared_matches,
    plan_copies,
    scan_files,
    unused_files,
)
from ..osutil import open_with_default_app

NAME = "gbcopy"
TITLE = "补料文件匹配拷贝"
SUMMARY = "按补料清单的「名称」列，在指定文件夹里找回 STP 与 PDF，拷到清单旁边"


def add_parser(subparsers: argparse._SubParsersAction) -> argparse.ArgumentParser:
    parser = subparsers.add_parser(
        NAME,
        help=SUMMARY,
        description=f"{TITLE}：{SUMMARY}。不带参数运行会依次弹窗选清单和搜索文件夹。",
        epilog='示例：speedup gbcopy "补料清单.xlsx" "D:\\WorkSpace\\订单" --dry-run',
    )
    parser.add_argument("excel", nargs="?", default=None, help="补料清单 .xlsx；省略则弹窗选择")
    parser.add_argument("search_dir", nargs="?", default=None, help="搜索文件夹；省略则弹窗选择")
    parser.add_argument(
        "--sheet",
        default=None,
        help=(
            "从哪张工作表读「名称」列。省略则自动：优先用名为「补料清单」的表，"
            "否则用第一张带「名称」列的表"
        ),
    )
    parser.add_argument(
        "--min-score",
        type=int,
        choices=(1, 3, 4),
        default=DEFAULT_MIN_SCORE,
        help="匹配强度：4=完全相同，3=允许版本号后缀（默认），1=宽松到互相包含",
    )
    parser.add_argument("--dry-run", action="store_true", help="只显示会拷什么，不真的拷贝")
    parser.add_argument("--no-open", action="store_true", help="结束后不要自动打开目标文件夹")
    return parser


def _choose_excel() -> str:
    chosen = ask_path(
        "file",
        "请输入清单文件路径：",
        partial(pick_file, "选择补料清单 Excel", FILTER_EXCEL),
        validate_file,
    )
    if not chosen:
        raise SpeedupError("用户未选择清单文件")
    return chosen


def _choose_search_dir(initial_dir: str) -> str:
    chosen = ask_path(
        "folder",
        "请输入搜索文件夹路径：",
        partial(pick_folder, "选择在哪个文件夹内搜索"),
        validate_folder,
        initial_dir=initial_dir,
    )
    if not chosen:
        raise SpeedupError("用户未选择搜索文件夹")
    return chosen


def render_match_table(results: list[MatchResult], search_root: str | None = None) -> str:
    """渲染匹配结果表；``!`` 表示需人工核对，``X`` 表示未找到。"""
    rows = []
    for index, result in enumerate(results, 1):
        rows.append(
            (
                str(index),
                result.name,
                _source_of(result, search_root),
                os.path.basename(result.model) if result.model else "-",
                os.path.basename(result.draw) if result.draw else "-",
                result.status,
            )
        )
    prefixes = [LEVEL_MARKERS.get(result.level, "  ") for result in results]
    return console.render_table(
        ("编号", "名称", "来源文件夹", "3D 文件", "图纸", "状态"),
        rows,
        max_widths=(4, 26, 24, 26, 26, 18),
        row_prefixes=prefixes,
    )


def _source_of(result: MatchResult, search_root: str | None) -> str:
    def relative(path: str) -> str:
        directory = os.path.dirname(path)
        if not search_root:
            return directory
        try:
            rel = os.path.relpath(directory, search_root)
        except ValueError:
            return directory
        return "." if rel == "." else rel

    model = relative(result.model) if result.model else ""
    draw = relative(result.draw) if result.draw else ""
    if model and draw and model != draw:
        return f"{model} | {draw}"
    return model or draw or "-"


def render_report(
    results: list[MatchResult],
    names: list[str],
    report: CopyReport,
    target_dir: str,
    *,
    dry_run: bool = False,
) -> str:
    """生成拷贝总结文本。"""
    not_found, lack = [], []
    for result in results:
        if not result.model and not result.draw:
            not_found.append(result.name)
        elif not result.model:
            lack.append((result.name, "3D 文件 (.stp/.step)"))
        elif not result.draw:
            lack.append((result.name, "图纸 (.pdf)"))

    lines: list[str] = []
    title = "匹 配 预 演（未拷贝）" if dry_run else "拷 贝 总 结"
    lines.append("=" * 78)
    lines.append(title)
    lines.append("=" * 78)
    lines.append(f"清单名称 {len(names)} 个        目标文件夹：{target_dir}")
    lines.append("")

    header = "【将拷贝】" if dry_run else "【已拷贝】"
    lines.append(f"{header}{len(report.copied)} 个")
    lines.extend(f"   +   {filename}      （{name}）" for name, filename in report.copied)
    if not report.copied:
        lines.append("   （无）")

    if report.skipped:
        lines.append("")
        lines.append(f"【同名未拷贝】{len(report.skipped)} 个 —— 目标文件夹已有同名文件，未覆盖")
        lines.extend(f"   =   {filename}      （{name}）" for name, filename in report.skipped)

    if report.inplace:
        lines.append("")
        lines.append(f"【本来就在目标文件夹】{len(report.inplace)} 个 —— 无需拷贝")
        lines.extend(f"   .   {filename}      （{name}）" for name, filename in report.inplace)

    if not_found or lack:
        lines.append("")
        lines.append(f"【未找到 / 缺失】{len(not_found) + len(lack)} 项")
        lines.extend(f"   X   {name}      3D 和 图纸 都没找到" for name in not_found)
        lines.extend(f"   X   {name}      缺 {what}" for name, what in lack)

    if report.failed:
        lines.append("")
        lines.append(f"【拷贝失败】{len(report.failed)} 个")
        lines.extend(f"   !   {filename}      {error}" for _name, filename, error in report.failed)

    lines.append("")
    lines.append(
        f"成功 {len(report.copied)} 个，同名跳过 {len(report.skipped)} 个，"
        f"本来就在 {len(report.inplace)} 个，完全没找到 {len(not_found)} 个名称，"
        f"另有 {len(lack)} 个只拷到一半，失败 {len(report.failed)} 个。"
    )
    lines.append("=" * 78)
    return "\n".join(lines)


def run(args: argparse.Namespace) -> int:
    console.ensure_safe_output()
    console.heading(TITLE)

    print("\n【步骤 1 / 4】确定补料清单")
    excel = args.excel or _choose_excel()
    excel = os.path.abspath(excel)
    if not os.path.isfile(excel):
        raise SpeedupError(f"清单文件不存在：{excel}")
    print(f"清单文件：{excel}")

    print("\n【步骤 2 / 4】读取「名称」列")
    extraction = extract_names(excel, sheet=args.sheet)
    print(f"  {extraction.summary()}")
    # 提醒必须打在名称列表**前面** —— 否则一屏 47 个名字刷过去，人根本反应不过来
    # 这里面混进了参考页的零件（这个坑踩过，见 AGENTS.md 的 P12）
    warning = extraction.warning()
    if warning:
        console.warn(warning)
    names = list(extraction.names)
    print(f"  共 {len(names)} 个零件：{'、'.join(names)}")

    target_dir = os.path.dirname(excel)
    print(f"  文件将拷贝到：{target_dir}")

    print("\n【步骤 3 / 4】确定搜索文件夹")
    search_dir = args.search_dir or _choose_search_dir(target_dir)
    search_dir = os.path.abspath(search_dir)
    if not os.path.isdir(search_dir):
        raise SpeedupError(f"搜索文件夹不存在：{search_dir}")
    print(f"搜索文件夹：{search_dir}")

    if os.path.normcase(search_dir) == os.path.normcase(target_dir):
        console.warn(
            "搜索文件夹和清单所在文件夹是同一个，子文件夹里匹配到的文件会被拷到清单这一层。"
        )

    print("\n正在扫描（含子文件夹）...")
    files = scan_files(search_dir, recursive=True)
    if not files:
        raise SpeedupError("该文件夹内没有找到任何 .stp/.step/.pdf 文件。")
    print(f"扫描完成，共 {len(files)} 个候选文件。")

    print(f"\n【步骤 4 / 4】匹配结果（匹配强度：{_score_label(args.min_score)}）")
    results = build_matches(names, files, min_score=args.min_score)
    print(render_match_table(results, search_root=search_dir))

    ok = sum(1 for r in results if r.level == "ok")
    fuzzy = sum(1 for r in results if r.level == "fuzzy")
    partial = sum(1 for r in results if r.level == "partial")
    miss = sum(1 for r in results if r.level == "miss")
    print(
        f"清单共 {len(names)} 个名称；完全/带版本匹配 {ok}，模糊匹配 {fuzzy}，"
        f"只找到一半 {partial}，未找到 {miss}"
    )
    print("（! 表示需要人工核对，X 表示未找到）")

    unused = unused_files(files, results)
    matched_count = len(files) - len(unused)
    print(f"\n扫描到 {len(files)} 个文件：{matched_count} 个与名称匹配，{len(unused)} 个无关（不会拷贝）。")

    clashes = find_shared_matches(results)
    if clashes:
        print("\n注意：以下文件被多个名称同时匹配到，请确认是否拷错：")
        for left, right, filename in clashes:
            print(f"    {left}  和  {right}  都匹配到  {filename}")

    actions = plan_copies(results, target_dir)
    report = execute_copies(actions) if not args.dry_run else _preview(actions)
    print()
    print(render_report(results, names, report, target_dir, dry_run=args.dry_run))

    if args.dry_run:
        console.info("这是 --dry-run，没有真的拷贝任何文件。去掉该参数即可执行。")
    elif os.path.isdir(target_dir) and not args.no_open:
        open_with_default_app(target_dir)
        print("\n已打开目标文件夹。")
    return 0


def _preview(actions: Sequence[CopyAction]) -> CopyReport:
    """``--dry-run`` 用的空跑版 :func:`execute_copies`：只分类，不落盘。"""
    report = CopyReport()
    for action in actions:
        filename = os.path.basename(action.dst)
        if action.kind == "inplace":
            report.inplace.append((action.name, filename))
        elif action.kind == "skip_same_name":
            report.skipped.append((action.name, filename))
        else:
            report.copied.append((action.name, filename))
    return report


def _score_label(min_score: int) -> str:
    return {
        4: "完全相同",
        3: "名称 + 版本号后缀",
        1: "宽松（互相包含也算）",
    }.get(min_score, f"自定义 {min_score}")
