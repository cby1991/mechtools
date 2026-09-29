"""文件 / 文件夹选择对话框 —— 纯 Python 实现，**彻底替代原来的 PowerShell 弹窗**。

为什么改
--------
原脚本用 ``subprocess`` 调 ``powershell -STA`` 去弹 WinForms 的
``OpenFileDialog`` / ``FolderBrowserDialog``，还借 ``OpenFileDialog`` 让用户
「选中文件夹里的任意一个 .stp 来间接选定文件夹」。这套做法有三个问题：

1. 依赖 PowerShell（本项目要求全 Python，不用 psh）；
2. 每次弹窗都要冷启动一个 PowerShell 进程，明显卡顿；
3. 临时 ``.ps1`` 写盘再删，杀软会误报，且多行脚本引号容易损坏。

现在统一走 :mod:`tkinter.filedialog`，零第三方依赖（tkinter 是标准库）。

分层回退
--------
``tkinter`` 缺失或在无图形界面的环境里运行时，:func:`pick_folder` /
:func:`pick_file` 返回 ``None``，由 :func:`ask_path` 自动回退到
「把路径拖进命令行窗口后回车」的控制台输入方式。所以工具永远不会因为
没有 GUI 而彻底不可用。

已知取舍
--------
tkinter 的目录选择框外观由 Tk / 系统决定，**不是**资源管理器同款大窗口。
这是纯 Python 方案下换取「零依赖 + 可单测」的代价，已在 README 中注明。
"""

from __future__ import annotations

import os
from collections.abc import Callable, Sequence
from functools import lru_cache
from typing import TypeAlias

__all__ = [
    "FileType",
    "dialog_available",
    "pick_folder",
    "pick_file",
    "ask_path",
    "validate_file",
    "validate_folder",
]

#: tkinter 的 (显示名, 通配符) 二元组；多个通配符用空格分隔
FileType: TypeAlias = tuple[str, str]

#: 常用过滤器，便于各工具复用
FILTER_STEP: tuple[FileType, ...] = (("STEP 模型", "*.stp *.step"), ("所有文件", "*.*"))
FILTER_PDF: tuple[FileType, ...] = (("PDF 图纸", "*.pdf"), ("所有文件", "*.*"))
FILTER_EXCEL: tuple[FileType, ...] = (
    ("Excel 工作簿", "*.xlsx *.xlsm"),
    ("Excel 97-2003", "*.xls"),
    ("所有文件", "*.*"),
)
FILTER_JSON: tuple[FileType, ...] = (("JSON 配置", "*.json"), ("所有文件", "*.*"))


@lru_cache(maxsize=1)
def dialog_available() -> bool:
    """tkinter 是否可用（只检查能否 import，不实际创建窗口）。"""
    try:
        import tkinter  # noqa: F401
    except Exception:  # ImportError / 平台相关的动态库缺失
        return False
    return True


def _run_dialog(callback: Callable[..., str]) -> str | None:
    """创建一个隐藏的置顶根窗口，跑完对话框后销毁。

    任何异常（含 ``TclError``：没有显示器、tk 初始化失败）都吞掉返回 ``None``，
    交给调用方回退到控制台输入 —— 不能因为弹窗失败就让工具挂掉。
    """
    try:
        import tkinter as tk
    except Exception:  # pragma: no cover - dialog_available() 已过滤
        return None

    root = None
    try:
        root = tk.Tk()
        root.withdraw()
        # 置顶，避免窗口被 IDE / 资源管理器挡住（原 PS 脚本也有同样的处理）
        try:
            root.attributes("-topmost", True)
        except tk.TclError:  # pragma: no cover - 少数窗口管理器不支持
            pass
        result = callback(root)
    except Exception:
        return None
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:  # pragma: no cover
                pass

    if not result:
        return None  # 用户点了取消
    return os.path.abspath(str(result))


def pick_folder(title: str, initial_dir: str | os.PathLike[str] | None = None) -> str | None:
    """弹窗选文件夹，返回绝对路径；取消或弹窗不可用返回 ``None``。"""
    if not dialog_available():
        return None

    def callback(root: object) -> str:
        from tkinter import filedialog

        options: dict[str, object] = {"parent": root, "title": title, "mustexist": True}
        if initial_dir and os.path.isdir(initial_dir):
            options["initialdir"] = str(initial_dir)
        return filedialog.askdirectory(**options)  # type: ignore[arg-type]

    return _run_dialog(callback)


def pick_file(
    title: str,
    filetypes: Sequence[FileType] | None = None,
    initial_dir: str | os.PathLike[str] | None = None,
) -> str | None:
    """弹窗选文件，返回绝对路径；取消或弹窗不可用返回 ``None``。"""
    if not dialog_available():
        return None

    def callback(root: object) -> str:
        from tkinter import filedialog

        options: dict[str, object] = {"parent": root, "title": title}
        if filetypes:
            options["filetypes"] = list(filetypes)
        if initial_dir and os.path.isdir(initial_dir):
            options["initialdir"] = str(initial_dir)
        return filedialog.askopenfilename(**options)  # type: ignore[arg-type]

    return _run_dialog(callback)


# ------------------------------------------------------------------ 手输回退


def validate_file(path: str | os.PathLike[str]) -> bool:
    return os.path.isfile(path)


def validate_folder(path: str | os.PathLike[str]) -> bool:
    return os.path.isdir(path)


def ask_path(
    kind: str,
    prompt: str,
    picker: Callable[..., str | None],
    validator: Callable[[str], bool],
    initial_dir: str | os.PathLike[str] | None = None,
    *,
    input_func: Callable[[str], str] = input,
) -> str | None:
    """先弹窗，失败则回退到控制台手输（支持把文件/文件夹拖进窗口）。

    :param kind: ``"file"`` 或 ``"folder"``，只影响回退时的提示文案
    :param picker: 弹窗函数，需接受 ``initial_dir`` 关键字参数
    :param validator: 校验路径是否可用
    :param input_func: 便于测试注入假的输入函数
    :returns: 可用路径；用户取消时返回 ``None``
    """
    if dialog_available():
        print("\n正在打开选择窗口...（若窗口被挡住，请查看任务栏）")
        chosen = picker(initial_dir=initial_dir)
        if chosen and validator(chosen):
            print(f"已选择：{chosen}")
            return chosen
        if chosen:
            print(f"选择的路径无效：{chosen}")
    else:
        print("\n[提示] 当前环境没有图形界面，改为手动输入路径。")

    noun = "文件" if kind == "file" else "文件夹"
    while True:
        print(f"\n（也可以直接把{noun}拖到本窗口后回车）")
        raw = input_func(prompt).strip().strip('"').strip("'")
        if not raw:
            print("已取消。")
            return None
        if validator(raw):
            return os.path.abspath(raw)
        print("路径不存在，请重新输入。")
