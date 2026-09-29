"""命令行入口：``speedup`` 默认开图形界面，也可以走子命令或文字菜单。

三种用法，同一套实现
--------------------
* **图形界面**：``uv run speedup`` —— 不带子命令时的默认行为（双击快捷方式也是这条路）。
  环境开不了窗口时自动降级到文字菜单，不会一上来就崩。
* **文字菜单**：``uv run speedup menu`` —— 列编号选工具，贴近原来双击 bat 的体验。
* **子命令**：``uv run speedup gbbuild "D:\\xxx" --no-open`` —— 便于脚本化与自动化测试。

三者不是三套逻辑：图形界面与菜单都复用 :func:`build_parser` 生成的默认参数去调同一个
``run()``，所以「界面上能跑」和「命令行能跑」永远一致。
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
from collections.abc import Callable, Sequence

from . import __version__, console
from .errors import SpeedupError, UserCancelled
from .tools import TOOLS, TOOL_MODULES

__all__ = ["build_parser", "main", "run_menu"]

PROGRAM = "speedup"
TAGLINE = "车间提速工具集 —— 全部纯 Python，不依赖 bat / PowerShell"

#: 菜单里除工具之外的固定项：``(按键, 标题, 说明)``
MENU_EXTRAS: tuple[tuple[str, str, str], ...] = (
    ("d", "环境体检", "检查 Python / 弹窗 / 编码 / 依赖是否就绪"),
    ("a", "历史脚本归档", "看看 archive/ 里留档的旧脚本"),
    ("0", "退出", ""),
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=PROGRAM,
        description=TAGLINE,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "不带任何参数运行会打开图形界面（开不了则自动降级为文字菜单）：\n"
            f"    uv run {PROGRAM}\n\n"
            "直接跑某个工具：\n"
            f"    uv run {PROGRAM} gbbuild \"D:\\WorkSpace\\订单\\0916\"\n"
        ),
    )
    parser.add_argument("--version", action="version", version=f"{PROGRAM} {__version__}")

    subparsers = parser.add_subparsers(dest="command", metavar="<子命令>")
    for module in TOOL_MODULES:
        module.add_parser(subparsers)
    subparsers.add_parser("gui", help="打开图形界面（不带参数运行时的默认行为）")
    subparsers.add_parser("menu", help="进入纯文字交互菜单（不开窗口）")
    doctor = subparsers.add_parser(
        "doctor", help="环境体检：Python / tkinter / 编码 / 依赖 / 工具链"
    )
    doctor.add_argument(
        "--install",
        action="store_true",
        help="发现必需依赖缺失时，由你主动触发补齐（会联网装包，大包可能要几分钟）",
    )
    subparsers.add_parser("archive", help="列出 archive/ 里归档的历史脚本")
    shortcut = subparsers.add_parser(
        "shortcut",
        help="创建「双击启动图形界面」的快捷方式（纯 Python，不写 bat/ps1）",
        description=(
            "在项目目录（或桌面）创建一个 .lnk 快捷方式，指向 pythonw.exe，"
            "双击即开图形界面且没有黑框控制台。"
        ),
    )
    shortcut.add_argument("--desktop", action="store_true", help="放到桌面，而不是项目目录")
    shortcut.add_argument("--dest", default=None, help="指定保存目录（优先级高于 --desktop）")
    shortcut.add_argument("--name", default=None, help="快捷方式显示名（默认「启动 speedup」）")
    return parser


# ------------------------------------------------------------------ 环境体检


def _probe(name: str, probe: Callable[[], str]) -> tuple[str, str]:
    try:
        return name, probe()
    except Exception as exc:  # 体检本身不能挂
        return name, f"不可用（{exc}）"


def _probe_python() -> str:
    return f"{sys.version.split()[0]}  ({sys.executable})"


def _probe_tkinter() -> str:
    from .dialogs import dialog_available

    if dialog_available():
        import tkinter

        return f"可用（Tk {tkinter.TkVersion}）—— 弹窗选择文件/文件夹正常"
    return "不可用 —— 会自动回退到「把路径拖进窗口后回车」的手输方式"


def _probe_openpyxl() -> str:
    import openpyxl

    return f"{openpyxl.__version__}（写采购清单、读清单都靠它）"


def _probe_pdfplumber() -> str:
    import pdfplumber

    return f"{getattr(pdfplumber, '__version__', 'unknown')}（读图纸：识别材料 + 提取表面处理要求）"


def _probe_pillow() -> str:
    from PIL import __version__ as pillow_version

    return f"{pillow_version}（Excel 嵌轴测图、裁白边靠它）"


def _probe_cadquery() -> str:
    """**可选**依赖：没装只影响「生成轴测图」，其它功能照常。

    所以探测结果里带「可选」两个字，``run_doctor`` 会据此**不把它算作失败**。
    """
    from .thumb import is_available

    if not is_available():
        return "未安装 —— 「生成轴测图」不可用（可选，装法：uv pip install cadquery）"
    try:
        import vtk  # noqa: F401  —— cadquery.vis 的渲染后端，缺了照样出不了图
    except ImportError:
        return "cadquery 在，但缺 vtk —— 渲染不可用（可选，装法：uv pip install vtk）"
    try:
        import cadquery

        return f"{cadquery.__version__}（生成零件轴测图）"
    except Exception as exc:  # noqa: BLE001 —— 体检本身不能挂
        return f"不可用（{exc}）"


def _has_module(name: str) -> bool:
    """这个模块能不能 import（用来判断 venv 里有没有 pip）。"""
    import importlib.util

    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError):
        return False


def _bundled_uv() -> str | None:
    """本机已知的 uv 绝对路径（uv 常常不在 PATH 上）。

    这是**这台机器**的兜底：本项目用 uv 建的 venv 不带 pip，
    而 uv 又常不在 PATH —— 所以要记一个已知位置。
    找不到就返回 None，让调用方去提示手装。
    """
    candidates = [
        os.path.expanduser(
            "~/.workbuddy/binaries/python/envs/default/Scripts/uv.exe"
        ),
        os.path.expanduser("~/.cargo/bin/uv.exe"),
        os.path.expanduser("~/.local/bin/uv"),
    ]
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    return None


def pick_installer() -> list[str] | None:
    """挑一个能装包的入口，返回**命令前缀**；一个都没有时返回 ``None``。

    按可靠性排序：

    1. 当前解释器自带的 ``python -m pip`` —— 最稳，装的必然是当前环境
    2. PATH 上的 ``uv pip`` —— 本项目用 uv 管理，这是常态
    3. 已知绝对路径的 ``uv`` —— uv 不在 PATH 时的兜底

    **为什么不能想当然用 ``python -m pip``**：uv 建的 venv 默认**不装 pip**，
    直接调会报 ``No module named pip``（实测踩过）。

    返回的只是**前缀**，``--python`` 由 :func:`install_command` 补 ——
    因为 uv 要求它出现在 ``install`` **之后**。
    """
    if _has_module("pip"):
        return [sys.executable, "-m", "pip"]

    found = shutil.which("uv")
    if found:
        return [found, "pip"]

    bundled = _bundled_uv()
    if bundled:
        return [bundled, "pip"]

    return None


def _is_uv(prefix: Sequence[str]) -> bool:
    """前缀是不是 uv（看可执行文件名）。"""
    return bool(prefix) and os.path.basename(prefix[0]).lower().startswith("uv")


def install_command(prefix: Sequence[str], *packages: str) -> list[str]:
    """拼出完整的装包命令。

    **uv 要额外钉住解释器，而且 ``--python`` 必须放在 ``install`` 之后**：

    * 对：``uv pip install --python <解释器> pillow``
    * 错：``uv pip --python <解释器> install pillow`` ← 实测直接报 Usage

    为什么非要钉住解释器：``uv pip install`` 默认装到「激活的那个环境」，
    而从图形界面或双击快捷方式启动时**根本没激活任何环境** —— 不指定就会
    装到别处，用户看到的是「明明装成功了却还是不可用」。
    """
    command = [*prefix, "install"]
    if _is_uv(prefix):
        command += ["--python", sys.executable]
    return [*command, *packages]


def run_doctor(  # noqa: PLR0912 —— 体检本来就分支多，拆开反而难读
    _args: argparse.Namespace | None = None, *, install: bool = False
) -> int:
    """打印环境体检结果。任何人反馈「跑不起来」时，先让他跑这个。

    :param install: ``True`` 时**由用户主动触发**补齐缺失的必需依赖。
        刻意不做成自动静默安装 —— 装包会改用户的环境，而且 cadquery
        这类大包要几分钟，静默装会让人以为程序卡死了。所以：**用户按了才装**，
        而且进度直接透传到终端。
    """
    console.ensure_safe_output()
    console.heading(f"{PROGRAM} {__version__} 环境体检")

    checks = [
        _probe("Python", _probe_python),
        _probe("tkinter 弹窗", _probe_tkinter),
        _probe("控制台编码", lambda: console.console_encoding()),
        _probe("openpyxl", _probe_openpyxl),
        _probe("pdfplumber", _probe_pdfplumber),
        _probe("pillow", _probe_pillow),
        _probe("cadquery（可选）", _probe_cadquery),
    ]
    console.print_table(("检查项", "结果"), checks, max_widths=(18, 58))

    encoding = console.console_encoding().lower()
    if encoding not in {"utf-8", "utf8", "cp936", "gbk", "gb2312"}:
        console.warn(
            f"控制台编码是 {encoding}，中文可能显示成问号。"
            "在中文 Windows 上应为 cp936；若在终端里，试试 chcp 936。"
        )
    else:
        console.ok(f"控制台编码 {encoding} 可以正常显示中文。")

    # 带「可选」标记的缺了不算失败 —— 它只影响某一个功能，不影响主流程（R29）
    missing = [
        name
        for name, result in checks
        if result.startswith("不可用") and "可选" not in name
    ]
    if not missing:
        console.ok("必需的依赖全部就绪。")
        return 0

    console.warn(f"以下必需依赖不可用：{'、'.join(missing)}")
    if install:
        return _install_missing(missing)

    console.warn("补齐办法（任选其一）：")
    print(f"        uv pip install {' '.join(_package_of(name) for name in missing)}")
    print(f"        pip install {' '.join(_package_of(name) for name in missing)}")
    return 1


def _package_of(probe_name: str) -> str:
    """体检项的名字 → PyPI 包名（两边不总是同名，所以显式映射）。"""
    return _PACKAGE_OF_PROBE.get(probe_name, probe_name)


#: 体检项 → PyPI 包名
_PACKAGE_OF_PROBE: dict[str, str] = {
    "openpyxl": "openpyxl",
    "pdfplumber": "pdfplumber",
    "pillow": "pillow",
    "cadquery（可选）": "cadquery",
}


def _install_missing(missing: Sequence[str]) -> int:
    """用户主动触发时装那些缺的包。进度直接透传到终端，不静默。"""
    installer = pick_installer()
    packages = [_package_of(name) for name in missing]

    if installer is None:
        console.warn("没找到可用的装包工具（venv 里没 pip，PATH 上也没 uv）。")
        console.warn("请手动执行：")
        print(f"        uv pip install {' '.join(packages)}")
        return 1

    import subprocess

    command = install_command(installer, *packages)
    console.section("开始安装")
    print(f"命令：{' '.join(command)}")
    print("（大包可能要几分钟，进度会实时刷出来；中途可以 Ctrl+C 中断）\n")

    try:
        code = subprocess.run(command, check=False).returncode  # noqa: S603
    except KeyboardInterrupt:
        console.warn("安装被中断。")
        return 1
    except OSError as exc:
        console.warn(f"装包命令起不来：{exc}")
        return 1

    if code != 0:
        console.warn(f"安装失败（退出码 {code}）。多半是网络问题，可手动重试。")
        return 1

    console.ok("安装完成 —— 下面重新体检一遍：")
    print()
    return run_doctor()


# ------------------------------------------------------------------ 归档清单


def archive_root() -> str:
    from .config import PROJECT_ROOT

    return os.path.join(PROJECT_ROOT, "archive")


def run_archive(_args: argparse.Namespace | None = None) -> int:
    """列出 ``archive/`` 下各分类与其中的脚本数量。"""
    from .config import PROJECT_ROOT

    root = os.path.join(PROJECT_ROOT, "archive")
    console.heading("历史脚本归档")

    if not os.path.isdir(root):
        console.warn(f"还没有归档目录：{root}")
        return 0

    rows: list[tuple[str, str, str]] = []
    for entry in sorted(os.listdir(root)):
        directory = os.path.join(root, entry)
        if not os.path.isdir(directory):
            continue
        # 必须递归：font_pipeline 这类分类下面还分了子目录
        files = [
            os.path.join(dirpath, name)
            for dirpath, _dirnames, filenames in os.walk(directory)
            for name in filenames
        ]
        scripts = [f for f in files if f.endswith(".py")]
        rows.append((entry, str(len(files)), f"{len(scripts)} 个 .py"))

    if not rows:
        console.warn("归档目录是空的。")
        return 0

    console.print_table(("分类", "文件数", "其中 Python"), rows, max_widths=(28, 8, 16))
    total = sum(int(count) for _name, count, _scripts in rows)
    print(f"\n合计 {total} 个文件（含各分类的 _PROVENANCE.md 溯源文件）。")
    print(f"归档根目录：{root}")
    print("这些是历史一次性脚本，**不参与构建、不写测试、不保证可运行**，仅作留档。")
    # 归档区只在本地（不进仓库），台账也跟着在本地，所以要判断存在再提示
    inventory = os.path.join(PROJECT_ROOT, "docs", "INVENTORY.md")
    if os.path.isfile(inventory):
        print(f"明细见：{inventory}")
    return 0


# ------------------------------------------------------------------ 交互菜单


def run_menu(
    *,
    input_func: Callable[[str], str] = input,
    pause: Callable[[], None] | None = None,
) -> int:
    """交互菜单主循环。返回值始终是 0（菜单本身不失败）。"""
    parser = build_parser()
    pause = pause or _pause

    while True:
        _print_menu()
        try:
            choice = input_func("请选择：").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return 0

        if choice in {"0", "q", "Q", "exit", "quit"}:
            print("已退出。")
            return 0

        if choice in {"d", "D"}:
            run_doctor()
            pause()
            continue
        if choice in {"a", "A"}:
            run_archive()
            pause()
            continue

        module = _resolve_choice(choice)
        if module is None:
            print(f"没有这个选项：{choice!r}，请输入列表里的编号。")
            continue

        args = parser.parse_args([module.NAME])
        try:
            module.run(args)
        except UserCancelled:
            print("\n已取消。")
        except SpeedupError as exc:
            console.fail(str(exc))
        except KeyboardInterrupt:
            print("\n已中断。")
        pause()


def _pause() -> None:
    try:
        input("\n按回车键返回菜单...")
    except (EOFError, KeyboardInterrupt):
        print()


def _resolve_choice(choice: str):
    """把菜单编号或子命令名解析成工具模块。"""
    ordered = sorted(TOOLS.values(), key=lambda module: module.NAME)
    if choice.isdigit():
        index = int(choice)
        if 1 <= index <= len(ordered):
            return ordered[index - 1]
        return None
    return TOOLS.get(choice)


def _print_menu() -> None:
    ordered = sorted(TOOLS.values(), key=lambda module: module.NAME)
    # 用 display_width 算宽度，不能用 f"{s:<12}" —— 那个按字符数补空格，中文标题会错位
    width = max(
        console.display_width(label)
        for label in (*[module.TITLE for module in ordered], *(item[1] for item in MENU_EXTRAS))
    )

    console.heading(f"{PROGRAM} {__version__}  ——  {TAGLINE}")
    print()
    for index, module in enumerate(ordered, 1):
        print(f"  {index}. {console.pad(module.TITLE, width)}  {module.SUMMARY}")
    print()
    for key, title, summary in MENU_EXTRAS:
        print(f"  {key:<2} {console.pad(title, width)}  {summary}".rstrip())
    print()


# ------------------------------------------------------------------ 快捷方式


def run_shortcut(args: argparse.Namespace | None = None) -> int:
    """创建「双击就能开图形界面」的快捷方式。

    这属于「装了什么东西到系统上」，所以先把要写什么原原本本打出来给人核对，
    再动手；而且**不覆盖已有文件时也不报错**，重复执行是安全的。
    """
    from .shortcut import (
        SHORTCUT_NAME,
        create_shortcut,
        default_destination,
        describe,
        plan_shortcut,
        shortcut_path,
    )

    console.heading("创建双击启动的快捷方式")

    spec = plan_shortcut(getattr(args, "name", None) or SHORTCUT_NAME)
    console.section("将要创建的快捷方式")
    print(describe(spec))

    destination = getattr(args, "dest", None)
    folder = destination if destination else default_destination(desktop=getattr(args, "desktop", False))
    path = shortcut_path(spec, folder)
    existed = path.exists()

    console.section("结果")
    created = create_shortcut(spec, path)
    console.ok(f"已{'覆盖' if existed else '创建'}：{created}")
    print()
    print("现在可以双击它启动图形界面（不会弹出黑框控制台）。")
    print("如果快捷方式上显示的还是 Python 默认图标，按 F5 刷新一下资源管理器即可。")
    if not existed:
        print("想放到桌面再加一个：uv run speedup shortcut --desktop")
    return 0


# ------------------------------------------------------------------ 图形界面


def _try_gui() -> tuple[int | None, str]:
    """尝试打开图形界面。

    :returns: ``(退出码, 失败原因)``。退出码为 ``None`` 表示**开不了**，
        失败原因是一句可以直接念给用户听的话；退出码不为 ``None`` 时原因恒为 ``""``。
    """
    try:
        from .gui import GuiUnavailable, launch
    except Exception as exc:  # tkinter 缺失、模块被改坏等
        return None, f"图形界面模块加载失败：{exc}"

    try:
        return launch(), ""
    except GuiUnavailable as exc:
        return None, str(exc)


def _alert_box(message: str) -> None:
    """``pythonw.exe`` 下没有控制台，图形界面又起不来时，只能弹系统对话框告知。"""
    if sys.platform != "win32":  # pragma: no cover - 非 Windows 走 print
        console.fail(message)
        return
    import ctypes

    try:
        ctypes.windll.user32.MessageBoxW(None, message, f"{PROGRAM} 启动失败", 0x10)
    except Exception:  # pragma: no cover
        pass


def run_gui() -> int:
    """``speedup gui``：强制开图形界面，开不了就明确报错、不降级。"""
    code, reason = _try_gui()
    if code is None:
        console.fail(reason)
        return 2
    return code


# ------------------------------------------------------------------ 入口


def main(argv: Sequence[str] | None = None) -> int:
    """进程入口。``argv`` 省略时取 ``sys.argv[1:]``。"""
    # pythonw.exe 下 sys.stdout 是 None，必须在任何 print 之前补上（见 console.ensure_streams）
    had_console = sys.stdout is not None
    console.ensure_streams()
    console.ensure_safe_output()

    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)

    command = getattr(args, "command", None)
    if command == "gui":
        return run_gui()
    if command == "menu":
        return run_menu()
    if command == "doctor":
        return run_doctor(args, install=getattr(args, "install", False))
    if command == "archive":
        return run_archive(args)
    if command == "shortcut":
        # 这个分支不是 TOOLS 里的工具（它不处理业务数据），单独兜一下写盘失败
        try:
            return run_shortcut(args)
        except SpeedupError as exc:
            console.fail(str(exc))
            return 2
    if command is None:
        code, reason = _try_gui()
        if code is not None:
            return code
        if not had_console:
            # 双击快捷方式但界面起不来：没有控制台可以打印，只能弹框
            _alert_box(f"{reason}\n\n请改用命令行排查：uv run {PROGRAM} menu")
            return 2
        console.warn(f"{reason}  已改用命令行菜单。")
        return run_menu()

    # 统一从 TOOLS 里按名字取模块再调 run()，而不是让 argparse 提前绑死函数对象。
    # 好处：菜单、命令行、单元测试走的是同一条路径，monkeypatch 也生效。
    module = TOOLS.get(command)
    if module is None:  # pragma: no cover - argparse 的 choices 已挡住
        parser.error(f"未知子命令：{command!r}")

    try:
        return int(module.run(args))
    except UserCancelled as exc:
        console.warn(str(exc) or "用户取消")
        return 2
    except SpeedupError as exc:
        console.fail(str(exc))
        return 2
    except KeyboardInterrupt:
        print("\n已中断。")
        return 130


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
