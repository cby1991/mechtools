"""图形界面（tkinter）—— 双击就能开，不用记命令行。

三种启动方式
------------
* 双击项目根目录的 **``启动 speedup.lnk``**（走 ``pythonw.exe``，没有黑框控制台）
* ``uv run speedup`` —— 不带子命令时**优先**开图形界面
* ``uv run speedup gui`` —— 强制开图形界面；环境不支持时明确报错，而不是静默降级

设计要点
--------
1. **复用同一套 ``run()``。** 界面只负责把控件里的值拼成 argv，再交给
   :func:`speedup.cli.build_parser` 解析成 ``Namespace``，最后调 ``module.run()``。
   所以「界面上能跑」和「命令行能跑」永远是同一份逻辑，不存在两套实现。
2. **工具跑在后台线程。** tkinter 的控件只能在主线程碰，所以工具在子线程执行，
   ``print`` 的输出被重定向进一个队列，主线程每 80ms 取一次、批量插进日志框。
   界面不会假死，长任务也能看到实时进度。
3. **不用在工作线程里弹窗。** 路径全部由界面先选好再传进去，工具内部那条
   「弹窗 → 手输」的回退路径在图形界面下不会被走到。
4. **日志用等宽字体（NSimSun）。** :mod:`speedup.console` 的表格是按
   「显示宽度」对齐的，换成比例字体就会错位。
"""

from __future__ import annotations

import contextlib
import io
import os
import queue
import sys
import threading
import time
import tkinter as tk
import traceback
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox, ttk
from typing import Any

from . import __version__, console
from .dialogs import (
    FILTER_EXCEL,
    dialog_available,
    pick_file,
    pick_folder,
)
from .errors import SpeedupError, UserCancelled
from .tools import TOOLS

__all__ = ["GuiUnavailable", "launch", "main"]


class GuiUnavailable(RuntimeError):
    """当前环境开不了图形界面（没有 tkinter，或没有可用的显示）。"""


# ------------------------------------------------------------------ 配色

NAVY = "#14365c"
NAVY_DARK = "#0e2745"
GOLD = "#e2a33d"
GOLD_DARK = "#c98a26"
BG = "#f4f6f9"
CARD = "#ffffff"
BORDER = "#d0d7e0"
TEXT = "#2b2b2b"
SUB = "#6b7785"
SOFT = "#e1e6ec"
SOFT_HOVER = "#d2dae4"
OK_GREEN = "#1f7a4d"
WARN_RED = "#b3261e"

#: 日志区必须是等宽字体，否则 console 的表格会错位。NSimSun(新宋体) 的 CJK 恰好占 2 列。
MONO_CANDIDATES: tuple[str, ...] = (
    "NSimSun",
    "新宋体",
    "SimSun",
    "宋体",
    "Consolas",
    "Courier New",
)
UI_CANDIDATES: tuple[str, ...] = (
    "Microsoft YaHei UI",
    "Microsoft YaHei",
    "微软雅黑",
    "SimHei",
    "黑体",
    "Segoe UI",
)


# ------------------------------------------------------------------ 参数表单描述


@dataclass(frozen=True)
class Field:
    """一个界面控件 ↔ 一个 argparse 参数的映射。

    :param dest: argparse 里的属性名（也是 ``Namespace`` 的字段名）
    :param label: 界面上显示的标签
    :param kind: ``folder`` / ``file`` / ``save`` / ``text`` / ``bool`` / ``choice``
    :param flag: 命令行开关，如 ``"--min-score"``；``None`` 表示位置参数
    :param required: 为空时是否挡住不跑
    :param default: 控件初始值
    :param hint: 控件下方的灰色说明
    :param filetypes: 文件选择框的过滤器
    :param choices: 下拉框的 ``(显示文本, 实际值)`` 列表
    :param negate: 勾选时**不加**该开关（用于 ``--no-open`` 这种反义参数）
    """

    dest: str
    label: str
    kind: str = "text"
    flag: str | None = None
    required: bool = False
    default: Any = None
    hint: str = ""
    filetypes: tuple[tuple[str, str], ...] | None = None
    choices: tuple[tuple[str, Any], ...] | None = None
    negate: bool = False


GBBUILD_FIELDS: tuple[Field, ...] = (
    Field(
        "folder",
        "目标文件夹",
        "folder",
        required=True,
        hint="连同子文件夹一起扫描 .stp / .step，并为每个模型找同名的 PDF 图纸",
    ),
    Field("out", "输出文件名", "text", flag="--out", hint="留空 = 用文件夹名 + .xlsx"),
    Field(
        "with_image",
        "生成轴测图",
        "bool",
        flag="--with-image",
        hint="给每个零件渲染一张等轴测图填进「图片」列（需先装 cadquery，每个零件约 5 秒）",
    ),
    Field(
        "no_open",
        "生成后自动打开表格",
        "bool",
        flag="--no-open",
        default=True,
        negate=True,
    ),
)

GBCOPY_FIELDS: tuple[Field, ...] = (
    Field(
        "excel",
        "采购清单",
        "file",
        required=True,
        filetypes=FILTER_EXCEL,
        hint="表格里要有「名称」列；文件会拷到这张清单所在的文件夹",
    ),
    Field(
        "sheet",
        "指定工作表",
        "text",
        flag="--sheet",
        hint="留空 = 自动（优先读名为「采购清单」的表；改名前的老文件叫「补料清单」，同样认）",
    ),
    Field(
        "search_dir",
        "搜索文件夹",
        "folder",
        required=True,
        hint="在哪个文件夹（含子文件夹）里找回 STP / PDF",
    ),
    Field(
        "min_score",
        "匹配强度",
        "choice",
        flag="--min-score",
        default=3,
        choices=(
            ("名称 + 版本号后缀（默认）", 3),
            ("完全相同", 4),
            ("宽松：互相包含也算", 1),
        ),
        hint="拿不准就保持默认；改宽松后结果里的 ! 需要人工核对",
    ),
    Field(
        "dry_run",
        "只预演，不真的拷贝",
        "bool",
        flag="--dry-run",
        default=False,
        hint="建议先勾上跑一次，确认匹配结果没问题再取消勾选真拷",
    ),
    Field(
        "no_open",
        "结束后自动打开目标文件夹",
        "bool",
        flag="--no-open",
        default=True,
        negate=True,
    ),
)

#: 每个工具的界面表单。新增工具时**必须**在这里补一份，否则它不会出现在图形界面里
#: （``tests/test_gui.py`` 里有断言盯着这件事）。
TOOL_FIELDS: dict[str, tuple[Field, ...]] = {
    "gbbuild": GBBUILD_FIELDS,
    "gbcopy": GBCOPY_FIELDS,
}


# ------------------------------------------------------------------ 输出重定向


class _QueueWriter(io.TextIOBase):
    """把 ``print()`` 的内容塞进队列，由主线程搬进日志框。

    ``encoding`` 故意返回 ``"utf-8"``：这个流不做任何编码转换，文本原样
    送到 Tk 控件里显示，所以「不会因编码丢字」这个事实用 utf-8 表达最贴切。
    顺带让 ``speedup doctor`` 里的编码检查在图形界面下不报假警告。
    """

    def __init__(self, sink: queue.Queue[tuple]) -> None:
        self._sink = sink

    @property
    def encoding(self) -> str:  # type: ignore[override]
        return "utf-8"

    def write(self, text: str) -> int:
        if text:
            self._sink.put(("log", text))
        return len(text)

    def flush(self) -> None:  # 队列没有缓冲，什么都不用做
        pass


# ------------------------------------------------------------------ DPI / 字体


def _enable_dpi_awareness() -> None:
    """必须在 ``tk.Tk()`` **之前**调用，否则高分屏上界面是糊的。"""
    if sys.platform != "win32":
        return
    import ctypes

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # System DPI aware
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()  # 老系统（Win8.1 以前）
    except Exception:
        pass


def _apply_scaling(root: tk.Tk) -> float:
    """按实际 DPI 调整 Tk 的字号缩放，返回一个像素缩放因子供布局使用。"""
    try:
        dpi = float(root.winfo_fpixels("1i"))
    except Exception:  # pragma: no cover - 极少见
        return 1.0
    factor = max(1.0, min(dpi / 96.0, 2.5))
    try:
        root.tk.call("tk", "scaling", dpi / 72.0)
    except Exception:  # pragma: no cover
        pass
    return factor


def _pick_family(root: tk.Tk, candidates: Sequence[str], fallback: str) -> str:
    try:
        available = set(tkfont.families(root))
    except Exception:  # pragma: no cover
        return fallback
    for name in candidates:
        if name in available:
            return name
    return fallback


# ------------------------------------------------------------------ 页面基类


class BasePage(tk.Frame):
    """一个页面 = 标题 + 说明 + 正文 + 一块运行日志。

    子类只需要实现 :meth:`build`（搭正文控件）和 :meth:`run_job`（返回一个
    可调用对象，在后台线程里执行）。不需要日志区的页面把 :attr:`HAS_LOG`
    设成 ``False`` 即可。
    """

    #: 子类接口
    title: str = ""
    summary: str = ""
    HAS_LOG: bool = True

    def __init__(self, master: tk.Misc, app: SpeedupApp):
        super().__init__(master, bg=BG)
        self.app = app
        self.log: tk.Text | None = None
        self._build_head()
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=18, pady=(0, 12))
        self.build(body)
        if self.HAS_LOG:
            self.log = self._build_log_area(body)

    # -- 子类接口 -------------------------------------------------------

    def build(self, parent: tk.Frame) -> None:
        """搭正文控件。默认什么都不放。"""

    def run_job(self) -> Callable[[], int]:
        """返回要在后台线程执行的可调用对象，返回值是退出码。"""
        raise NotImplementedError

    def can_run(self) -> str | None:
        """返回 ``None`` 表示可以跑；返回字符串表示挡下来的原因。"""
        return None

    # -- 公共结构 -------------------------------------------------------

    def _build_head(self) -> None:
        head = tk.Frame(self, bg=BG)
        head.pack(fill="x", padx=18, pady=(16, 10))
        tk.Label(
            head,
            text=self.title,
            bg=BG,
            fg=NAVY,
            font=self.app.font_title,
            anchor="w",
        ).pack(fill="x")
        if self.summary:
            tk.Label(
                head,
                text=self.summary,
                bg=BG,
                fg=SUB,
                font=self.app.font_small,
                anchor="w",
                justify="left",
                wraplength=self.app.wrap_px,
            ).pack(fill="x", pady=(4, 0))

    def _build_log_area(self, parent: tk.Frame) -> tk.Text:
        card = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True, pady=(12, 0))

        bar = tk.Frame(card, bg=CARD)
        bar.pack(fill="x", padx=10, pady=(8, 4))
        tk.Label(
            bar, text="运行日志", bg=CARD, fg=NAVY, font=self.app.font_bold, anchor="w"
        ).pack(side="left")
        self.app.small_button(bar, "复制日志", self.copy_log).pack(side="right")
        self.app.small_button(bar, "清空日志", self.clear_log).pack(
            side="right", padx=(0, 6)
        )

        wrap = tk.Frame(card, bg=CARD)
        wrap.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        # wrap="none"：console 渲染的表格不能被自动折行，否则对齐全废，改用横向滚动条
        log = tk.Text(
            wrap,
            wrap="none",
            height=12,
            bg="#fbfcfd",
            fg=TEXT,
            font=self.app.font_mono,
            relief="flat",
            highlightthickness=0,
            padx=8,
            pady=6,
            state="disabled",
        )
        vbar = ttk.Scrollbar(wrap, orient="vertical", command=log.yview)
        hbar = ttk.Scrollbar(wrap, orient="horizontal", command=log.xview)
        log.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        log.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        return log

    # -- 日志操作 -------------------------------------------------------

    def append_log(self, text: str) -> None:
        if self.log is None:  # pragma: no cover - 只有 HAS_LOG=False 的页面会走到
            return
        self.log.configure(state="normal")
        self.log.insert("end", text)
        # 日志太长时砍掉开头，避免长时间运行把内存吃掉
        if int(self.log.index("end-1c").split(".")[0]) > 5000:
            self.log.delete("1.0", "500.0")
        self.log.see("end")
        self.log.configure(state="disabled")

    def clear_log(self) -> None:
        if self.log is None:
            return
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def copy_log(self) -> None:
        if self.log is None:
            return
        content = self.log.get("1.0", "end-1c")
        if not content.strip():
            return
        self.app.root.clipboard_clear()
        self.app.root.clipboard_append(content)
        self.app.set_status("日志已复制到剪贴板。")


# ------------------------------------------------------------------ 工具页


class ToolPage(BasePage):
    """一个工具一页：上面是参数表单，下面是运行日志。"""

    def __init__(
        self,
        master: tk.Misc,
        app: SpeedupApp,
        module: Any,
        fields: tuple[Field, ...],
    ):
        self.module = module
        self.fields = fields
        self.title = module.TITLE
        self.summary = module.SUMMARY
        self._vars: dict[str, tk.Variable] = {}
        self._widgets: dict[str, tk.Widget] = {}
        super().__init__(master, app)

    def build(self, parent: tk.Frame) -> None:
        card = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x")
        form = tk.Frame(card, bg=CARD)
        form.pack(fill="x", padx=4, pady=6)
        form.columnconfigure(1, weight=1)

        row = 0
        for field in self.fields:
            row = self._build_field(form, row, field)

        buttons = tk.Frame(card, bg=CARD)
        buttons.pack(fill="x", padx=14, pady=(2, 12))
        self.run_button = self.app.primary_button(
            buttons, f"开始执行「{self.title}」", self.on_run
        )
        self.run_button.pack(side="left")
        self.app.small_button(buttons, "重置表单", self.reset).pack(side="left", padx=(10, 0))

    def _build_field(self, form: tk.Frame, row: int, field: Field) -> int:
        tk.Label(
            form,
            text=field.label,
            bg=CARD,
            fg=TEXT,
            font=self.app.font_body,
            anchor="w",
        ).grid(row=row, column=0, sticky="w", padx=(14, 12), pady=(10, 0))
        row += 1

        if field.kind == "bool":
            var = tk.BooleanVar(master=form, value=bool(field.default))
            self._vars[field.dest] = var
            widget = tk.Checkbutton(
                form,
                variable=var,
                bg=CARD,
                fg=TEXT,
                activebackground=CARD,
                activeforeground=TEXT,
                selectcolor=CARD,
                font=self.app.font_body,
                anchor="w",
                highlightthickness=0,
                bd=0,
                cursor="hand2",
            )
            widget.grid(row=row - 1, column=1, columnspan=2, sticky="w", pady=(10, 0))
            self._widgets[field.dest] = widget
            return self._hint(form, row, field)

        if field.kind == "choice":
            var = tk.StringVar(master=form, value=str(field.default))
            self._vars[field.dest] = var
            choices = field.choices or ()
            labels = [label for label, _value in choices]
            widget = ttk.Combobox(
                form,
                textvariable=var,  # 必须接上：否则 current() 只改显示、不改值
                values=labels,
                state="readonly",
                font=self.app.font_body,
                width=28,
            )
            for index, (_label, value) in enumerate(choices):
                if value == field.default:
                    widget.current(index)
                    break
            else:  # 没有和默认值匹配的项，退到第一项
                if labels:
                    widget.current(0)
            widget.grid(row=row - 1, column=1, columnspan=2, sticky="w", pady=(10, 0))
            self._widgets[field.dest] = widget
            return self._hint(form, row, field)

        var = tk.StringVar(master=form, value="" if field.default is None else str(field.default))
        self._vars[field.dest] = var
        entry = tk.Entry(
            form,
            textvariable=var,
            font=self.app.font_body,
            bg="#ffffff",
            fg=TEXT,
            relief="solid",
            bd=1,
            highlightthickness=0,
        )
        entry.grid(row=row - 1, column=1, sticky="ew", pady=(10, 0))
        self._widgets[field.dest] = entry

        if field.kind in {"folder", "file", "save"}:
            self.app.small_button(
                form, "浏览...", lambda f=field: self._browse(f)
            ).grid(row=row - 1, column=2, sticky="w", padx=(8, 14), pady=(10, 0))

        return self._hint(form, row, field)

    def _hint(self, form: tk.Frame, row: int, field: Field) -> int:
        """灰色说明文字占一整行；返回下一个可用的行号。"""
        if not field.hint:
            return row
        tk.Label(
            form,
            text=field.hint,
            bg=CARD,
            fg=SUB,
            font=self.app.font_small,
            anchor="w",
            justify="left",
            wraplength=self.app.hint_px,
        ).grid(row=row, column=1, columnspan=2, sticky="w", pady=(1, 0))
        return row + 1

    # -- 取值 / 组 argv --------------------------------------------------

    def _browse(self, field: Field) -> None:
        current = str(self._vars[field.dest].get()).strip()
        initial = os.path.dirname(current) if current else None
        if field.kind == "folder":
            chosen = pick_folder(f"选择{field.label}", initial_dir=initial)
        elif field.kind == "save":
            from tkinter import filedialog

            chosen = filedialog.asksaveasfilename(
                title=f"选择{field.label}",
                filetypes=list(field.filetypes or ()),
                defaultextension=os.path.splitext(current)[1] or ".png",
                initialdir=initial or None,
            )
        else:
            chosen = pick_file(f"选择{field.label}", field.filetypes, initial_dir=initial)
        if chosen:
            self._vars[field.dest].set(chosen)

    def _value_of(self, field: Field) -> Any:
        raw = self._vars[field.dest].get()
        if field.kind == "bool":
            return bool(raw)
        if field.kind == "choice":
            for label, value in field.choices or ():
                if label == raw:
                    return value
            return field.default
        return str(raw).strip()

    def field(self, dest: str) -> Field:
        """按 argparse 的属性名找字段定义。"""
        for field in self.fields:
            if field.dest == dest:
                return field
        raise KeyError(f"{self.module.NAME} 没有名为 {dest!r} 的字段")

    def set_value(self, dest: str, value: Any) -> None:
        """以编程方式设置字段值（不想去摸控件变量时用）。

        下拉框可以直接给**语义值**（如 ``min_score=1``）也可以给显示文本。
        """
        spec = self.field(dest)
        if spec.kind == "choice":
            widget = self._widgets[dest]
            for index, (label, choice_value) in enumerate(spec.choices or ()):
                if choice_value == value or label == value:
                    widget.current(index)  # type: ignore[attr-defined]
                    return
            raise ValueError(f"{dest} 没有这个选项：{value!r}")
        self._vars[dest].set(value)

    def value_of(self, dest: str) -> Any:
        """读取字段当前值的语义化结果（下拉框返回的是值，不是显示文本）。"""
        return self._value_of(self.field(dest))

    def build_argv(self) -> list[str]:
        """把表单里的值拼成 argv 片段（位置参数在前，开关在后）。

        :raises ValueError: 必填项为空
        """
        positionals: list[str] = []
        options: list[str] = []
        for field in self.fields:
            value = self._value_of(field)
            if field.kind == "bool":
                enabled = bool(value)
                if field.negate:
                    enabled = not enabled
                if enabled and field.flag:
                    options.append(field.flag)
                continue
            if field.kind == "choice":
                if field.flag:
                    options += [field.flag, str(value)]
                continue
            text = str(value or "").strip()
            if not text:
                if field.required:
                    raise ValueError(f"「{field.label}」还没选。{field.hint}".strip())
                continue
            if field.flag:
                options += [field.flag, text]
            else:
                positionals.append(text)
        return positionals + options

    def reset(self) -> None:
        for field in self.fields:
            if field.kind == "bool":
                self._vars[field.dest].set(bool(field.default))
            elif field.kind == "choice":
                for index, (label, _value) in enumerate(field.choices or ()):
                    if _value == field.default:
                        self._widgets[field.dest].current(index)  # type: ignore[attr-defined]
                        break
            else:
                self._vars[field.dest].set("")
        self.app.set_status("表单已重置。")

    # -- 运行 -----------------------------------------------------------

    def can_run(self) -> str | None:
        try:
            self.build_argv()
        except ValueError as exc:
            return str(exc)
        return None

    def on_run(self) -> None:
        if self.app.busy:
            return
        problem = self.can_run()
        if problem:
            messagebox.showwarning("还差一点", problem, parent=self.app.root)
            return
        argv = self.build_argv()
        self.clear_log()
        self.app.set_status(f"正在执行「{self.title}」...")
        self.app.run_task(self, lambda: self._invoke(argv))

    def _invoke(self, argv: list[str]) -> int:
        # 关键：走 build_parser，保证拿到的 Namespace 和命令行完全一致
        from .cli import build_parser

        args = build_parser().parse_args([self.module.NAME, *argv])
        return int(self.module.run(args))

    def set_running(self, running: bool) -> None:
        self.run_button.configure(state="disabled" if running else "normal")


# ------------------------------------------------------------------ 动作页（体检 / 归档）


def _doctor_job(install: bool = False) -> int:
    """跑体检。``install=True`` 时补齐缺失的必需依赖（由用户点按钮触发）。"""
    from .cli import run_doctor

    return run_doctor(install=install)


def _archive_job() -> int:
    from .cli import run_archive

    return run_archive()


class ActionPage(BasePage):
    """没有参数、点一下就跑的页面（环境体检、归档清单）。"""

    def __init__(
        self,
        master: tk.Misc,
        app: SpeedupApp,
        title: str,
        summary: str,
        button_text: str,
        job: Callable[[], int],
        intro: str = "",
        extra: tuple[str, Callable[[], int]] | None = None,
    ):
        self.title = title
        self.summary = summary
        self._button_text = button_text
        self._job = job
        self._intro = intro
        #: 可选的第二个按钮 ``(文字, 任务)`` —— 比如体检页的「补齐依赖」
        self._extra = extra
        self.extra_button: tk.Button | None = None
        super().__init__(master, app)

    def build(self, parent: tk.Frame) -> None:
        card = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x")
        inner = tk.Frame(card, bg=CARD)
        inner.pack(fill="x", padx=14, pady=12)
        if self._intro:
            tk.Label(
                inner,
                text=self._intro,
                bg=CARD,
                fg=TEXT,
                font=self.app.font_body,
                anchor="w",
                justify="left",
                wraplength=self.app.hint_px,
            ).pack(fill="x", pady=(0, 10))
        self.run_button = self.app.primary_button(inner, self._button_text, self.on_run)
        self.run_button.pack(side="left")
        if self._extra is not None:
            label, job = self._extra
            self.extra_button = self.app.small_button(
                inner, label, lambda: self._on_extra(label, job)
            )
            self.extra_button.pack(side="left", padx=(6, 0))

    def on_run(self) -> None:
        if self.app.busy:
            return
        self.clear_log()
        self.app.set_status(f"正在执行「{self.title}」...")
        self.app.run_task(self, self._job)

    def _on_extra(self, label: str, job: Callable[[], int]) -> None:
        """第二个按钮走和主按钮**同一条路径**（R19：界面只负责拼 argv / 派任务）。"""
        if self.app.busy:
            return
        self.clear_log()
        self.app.set_status(f"正在执行「{label}」...")
        self.app.run_task(self, job)

    def run_job(self) -> Callable[[], int]:
        return self._job

    def set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.run_button.configure(state=state)
        if self.extra_button is not None:
            self.extra_button.configure(state=state)


# ------------------------------------------------------------------ 说明页


HELP_TEXT = """\
这个窗口是 speedup 的图形界面 —— 不用记任何命令。

────────────────────────────────────────────────────────────
一、两个工具怎么用
────────────────────────────────────────────────────────────
1. 生成采购清单
   选一个装 STP 模型的文件夹 → 点执行。
   程序会扫出所有 .stp/.step，为每个模型找同目录同名的 PDF 图纸，
   从图纸标题栏里读出材料，最后生成一份 Excel 采购清单。
   图纸没找到、或者图纸里没写材料，会在日志里列出来（材料列留空，需要手填）。

2. 采购文件匹配拷贝
   选 Excel 采购清单 + 一个搜索文件夹 → 点执行。
   按清单「名称」列去搜索文件夹里把对应的 STP 和 PDF 找出来，拷到清单旁边。
   【重要】建议先勾上「只预演，不真的拷贝」，确认匹配结果对得上，再取消勾选真拷。
   同名文件永远不会被覆盖。
   清单里有多张工作表时，只会读一张（优先读名为「采购清单」的那张），
   日志里会说明读了哪张、忽略了哪张。

────────────────────────────────────────────────────────────
二、参数不知道填什么？
────────────────────────────────────────────────────────────
· 每个输入框下面都有一行灰字说明，先看那个。
· 路径类都不用手打，点「浏览...」选就行。
· 拿不准的参数保持默认值即可。

────────────────────────────────────────────────────────────
三、还想用命令行？
────────────────────────────────────────────────────────────
在项目目录执行（界面上做的每一件事都有等价的命令）：
    uv run speedup                 图形界面（就是现在这个窗口）
    uv run speedup menu            纯文字菜单，不开窗口
    uv run speedup doctor          环境体检
    uv run speedup gbbuild "D:\\某文件夹" --no-open
    uv run speedup gbcopy "清单.xlsx" "D:\\搜索目录" --dry-run

────────────────────────────────────────────────────────────
四、窗口没反应 / 报错怎么办？
────────────────────────────────────────────────────────────
1. 工具是在后台线程跑的，界面不会卡；日志是一行一行实时刷出来的。
2. 真报错了，日志里会有 [失败] 或 [异常] 开头的段落 ——
   点右上角「复制日志」，把内容发给小埃，就能定位。
3. 环境有问题时，先跑一下「环境体检」页。
"""


class HelpPage(BasePage):
    title = "使用说明"
    summary = "两个工具怎么用、参数怎么填、出问题怎么办"
    HAS_LOG = False

    #: 只读的正文控件（``build`` 里赋值）
    text: tk.Text | None = None

    def build(self, parent: tk.Frame) -> None:
        card = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True)
        wrap = tk.Frame(card, bg=CARD)
        wrap.pack(fill="both", expand=True, padx=12, pady=12)
        text = tk.Text(
            wrap,
            wrap="word",
            bg=CARD,
            fg=TEXT,
            font=self.app.font_body,
            relief="flat",
            highlightthickness=0,
            padx=4,
            pady=2,
        )
        bar = ttk.Scrollbar(wrap, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=bar.set)
        text.insert("1.0", HELP_TEXT)
        text.configure(state="disabled")
        text.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")
        self.text = text

    def run_job(self) -> Callable[[], int]:  # pragma: no cover - 说明页不跑任务
        return lambda: 0


# ------------------------------------------------------------------ 主窗口


class SpeedupApp:
    """主窗口：左边导航，右边页面，底下一条状态栏。"""

    def __init__(self, root: tk.Tk, factor: float = 1.0):
        self.root = root
        self.factor = factor
        self.busy = False
        self.queue: queue.Queue[tuple] = queue.Queue()
        self.active_page: BasePage | None = None
        self.started_at = 0.0
        self._writer = _QueueWriter(self.queue)

        self._setup_window()
        self._setup_fonts()
        self._setup_style()
        self._build_layout()
        self.pages = self._build_pages()
        self._build_nav()
        self.show(0)
        self._drain_id: str | None = self.root.after(80, self._drain)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # -- 初始化 ---------------------------------------------------------

    def _setup_window(self) -> None:
        self.root.title(f"speedup {__version__} —— 车间提速工具集")
        self.root.configure(bg=BG)
        self._apply_icon()

        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        width = min(int(980 * self.factor), int(screen_w * 0.92))
        height = min(int(760 * self.factor), int(screen_h * 0.86))
        x = max((screen_w - width) // 2, 0)
        y = max((screen_h - height) // 3, 0)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(min(int(840 * self.factor), width), min(int(560 * self.factor), height))

        # 先置顶再放开，避免窗口被 IDE / 资源管理器挡住（和 dialogs.py 同一套路）
        try:
            self.root.attributes("-topmost", True)
            self.root.after(400, lambda: self.root.attributes("-topmost", False))
        except tk.TclError:  # pragma: no cover
            pass
        self.root.lift()

    def _apply_icon(self) -> None:
        icon = Path(__file__).resolve().parent / "assets" / "speedup.ico"
        if not icon.is_file():
            return
        try:
            self.root.iconbitmap(default=str(icon))
        except tk.TclError:  # pragma: no cover - 非 Windows 或缺图标支持
            pass

    def _setup_fonts(self) -> None:
        ui = _pick_family(self.root, UI_CANDIDATES, "TkDefaultFont")
        mono = _pick_family(self.root, MONO_CANDIDATES, "TkFixedFont")
        self.font_title = (ui, 15, "bold")
        self.font_bold = (ui, 10, "bold")
        self.font_body = (ui, 10)
        self.font_small = (ui, 9)
        self.font_mono = (mono, 9)
        self.font_nav = (ui, 10)
        # 需要像素宽度的地方（换行、导航栏）随 DPI 放大
        self.wrap_px = int(720 * self.factor)
        self.hint_px = int(660 * self.factor)
        self.nav_px = int(190 * self.factor)

    def _setup_style(self) -> None:
        style = ttk.Style(self.root)
        for theme in ("clam", "vista", "default"):
            try:
                style.theme_use(theme)
                break
            except tk.TclError:  # pragma: no cover
                continue
        style.configure(
            "TCombobox",
            fieldbackground="#ffffff",
            background="#ffffff",
            foreground=TEXT,
            arrowcolor=NAVY,
        )
        style.map("TCombobox", fieldbackground=[("readonly", "#ffffff")])
        style.configure("TScrollbar", background=SOFT, troughcolor=BG, bordercolor=BG)

    def _build_layout(self) -> None:
        # 所有东西都挂在这一层下面，退出时一次性销毁干净
        self.root_frame = tk.Frame(self.root, bg=BG)
        self.root_frame.pack(fill="both", expand=True)

        header = tk.Frame(self.root_frame, bg=NAVY)
        header.pack(fill="x")
        tk.Label(
            header,
            text=f"speedup  {__version__}",
            bg=NAVY,
            fg="#ffffff",
            font=(self.font_title[0], 14, "bold"),
        ).pack(anchor="w", padx=18, pady=(12, 0))
        tk.Label(
            header,
            text="车间提速工具集 —— 全部纯 Python，不依赖 bat / PowerShell",
            bg=NAVY,
            fg="#c8d6e6",
            font=self.font_small,
        ).pack(anchor="w", padx=18, pady=(2, 12))

        self.status = tk.Label(
            self.root_frame,
            text="就绪。左侧选一个工具，填好参数点执行。",
            bg=SOFT,
            fg=TEXT,
            font=self.font_small,
            anchor="w",
            padx=14,
            pady=5,
        )
        self.status.pack(fill="x", side="bottom")

        middle = tk.Frame(self.root_frame, bg=BG)
        middle.pack(fill="both", expand=True)

        self.nav = tk.Frame(middle, bg=SOFT, width=self.nav_px)
        self.nav.pack(side="left", fill="y")
        self.nav.pack_propagate(False)

        self.container = tk.Frame(middle, bg=BG)
        self.container.pack(side="left", fill="both", expand=True)

    def _build_pages(self) -> list[BasePage]:
        pages: list[BasePage] = []
        for name in sorted(TOOLS):
            fields = TOOL_FIELDS.get(name)
            if fields is None:  # 没配表单的工具不进图形界面（测试里会拦住）
                continue
            pages.append(ToolPage(self.container, self, TOOLS[name], fields))
        pages.append(
            ActionPage(
                self.container,
                self,
                "环境体检",
                "检查 Python / tkinter / 控制台编码 / 必需依赖是否就绪（含可选出图能力）",
                "开始体检",
                lambda: _doctor_job(),
                intro=(
                    "跑不起来任何工具时，先点这里。缺少依赖会明确告诉你是哪一个。\n"
                    "「补齐依赖」由你点才动手 —— 要联网装包，大包可能要几分钟。"
                ),
                extra=("补齐依赖", lambda: _doctor_job(install=True)),
            )
        )
        pages.append(
            ActionPage(
                self.container,
                self,
                "历史脚本归档",
                "看看 archive/ 里留档的旧脚本有多少、都分在哪几类",
                "查看归档",
                _archive_job,
                intro="这些是历史上一次性的旧脚本，只做留档，不参与构建、不保证能跑。",
            )
        )
        pages.append(HelpPage(self.container, self))
        return pages

    def _build_nav(self) -> None:
        tk.Label(
            self.nav,
            text="工具",
            bg=SOFT,
            fg=SUB,
            font=self.font_small,
            anchor="w",
            padx=14,
        ).pack(fill="x", pady=(12, 4))
        self.nav_buttons: list[tk.Button] = []
        for index, page in enumerate(self.pages):
            if index == 3:
                tk.Frame(self.nav, bg=BORDER, height=1).pack(fill="x", pady=(10, 0))
                tk.Label(
                    self.nav,
                    text="其他",
                    bg=SOFT,
                    fg=SUB,
                    font=self.font_small,
                    anchor="w",
                    padx=14,
                ).pack(fill="x", pady=(8, 4))
            button = tk.Button(
                self.nav,
                text=page.title,
                anchor="w",
                relief="flat",
                bd=0,
                highlightthickness=0,
                font=self.font_nav,
                cursor="hand2",
                padx=14,
                pady=8,
                command=lambda i=index: self.show(i),
            )
            button.pack(fill="x")
            button.bind("<Enter>", lambda _e, i=index: self._hover(i, True))
            button.bind("<Leave>", lambda _e, i=index: self._hover(i, False))
            self.nav_buttons.append(button)

    # -- 控件工厂 -------------------------------------------------------

    def primary_button(self, parent: tk.Misc, text: str, command: Callable[[], None]) -> tk.Button:
        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=NAVY,
            fg="#ffffff",
            activebackground=NAVY_DARK,
            activeforeground="#ffffff",
            disabledforeground="#b9c4d2",
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=self.font_bold,
            cursor="hand2",
            padx=16,
            pady=7,
        )
        button.bind("<Enter>", lambda _e: self._tint(button, NAVY_DARK))
        button.bind("<Leave>", lambda _e: self._tint(button, NAVY))
        return button

    def small_button(self, parent: tk.Misc, text: str, command: Callable[[], None]) -> tk.Button:
        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=SOFT,
            fg=TEXT,
            activebackground=SOFT_HOVER,
            activeforeground=TEXT,
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=self.font_small,
            cursor="hand2",
            padx=10,
            pady=4,
        )
        button.bind("<Enter>", lambda _e: self._tint(button, SOFT_HOVER))
        button.bind("<Leave>", lambda _e: self._tint(button, SOFT))
        return button

    @staticmethod
    def _tint(button: tk.Button, colour: str) -> None:
        if str(button.cget("state")) != "disabled":
            button.configure(bg=colour)

    def _hover(self, index: int, entering: bool) -> None:
        if index == getattr(self, "_current", -1):
            return  # 当前页始终保持高亮，不做悬停变色
        self.nav_buttons[index].configure(bg=SOFT_HOVER if entering else SOFT)

    def _highlight(self, index: int) -> None:
        """把导航栏里第 ``index`` 项刷成高亮，其余恢复常态。"""
        for position, button in enumerate(self.nav_buttons):
            active = position == index
            button.configure(
                bg=NAVY if active else SOFT,
                fg="#ffffff" if active else TEXT,
                font=(self.font_nav[0], self.font_nav[1], "bold" if active else "normal"),
            )
        self._current = index

    # -- 页面切换 -------------------------------------------------------

    def show(self, index: int) -> None:
        if self.busy:
            self.set_status("任务正在跑，等它结束再切换页面。")
            return
        for page in self.pages:
            page.pack_forget()
        self.pages[index].pack(fill="both", expand=True)
        self._highlight(index)
        self.set_status(self.pages[index].summary or "就绪。")

    # -- 任务调度 -------------------------------------------------------

    def run_task(self, page: BasePage, job: Callable[[], int]) -> None:
        self.busy = True
        self.active_page = page
        self.started_at = time.monotonic()
        page.clear_log()
        page.set_running(True)
        for button in self.nav_buttons:
            button.configure(state="disabled")
        self.set_status(f"正在执行「{page.title}」...")
        threading.Thread(
            target=self._worker,
            args=(job,),
            name="speedup-task",
            daemon=True,
        ).start()

    def _worker(self, job: Callable[[], int]) -> None:
        try:
            with contextlib.redirect_stdout(self._writer), contextlib.redirect_stderr(
                self._writer
            ):
                code = job()
            self.queue.put(("done", int(code), "ok", time.monotonic() - self.started_at))
        except UserCancelled as exc:
            self.queue.put(("log", f"\n[取消] {exc or '用户取消了操作'}\n"))
            self.queue.put(("done", 2, "cancel", time.monotonic() - self.started_at))
        except SpeedupError as exc:
            self.queue.put(("log", f"\n[失败] {exc}\n"))
            self.queue.put(("done", 2, "fail", time.monotonic() - self.started_at))
        except SystemExit as exc:
            code = exc.code if isinstance(exc.code, int) else 2
            self.queue.put(("log", f"\n[退出] 参数解析失败（退出码 {code}）。\n"))
            self.queue.put(("done", code, "fail", time.monotonic() - self.started_at))
        except KeyboardInterrupt:
            self.queue.put(("log", "\n[中断] 已取消。\n"))
            self.queue.put(("done", 130, "cancel", time.monotonic() - self.started_at))
        except Exception:
            self.queue.put(
                (
                    "log",
                    "\n[异常] 这是没预料到的错误。点「复制日志」把下面这段发给小埃：\n"
                    + traceback.format_exc()
                    + "\n",
                )
            )
            self.queue.put(("done", 1, "error", time.monotonic() - self.started_at))

    def _drain(self) -> None:
        """主线程每 80ms 把队列里的日志批量搬进日志框。"""
        chunks: list[str] = []
        finished: tuple | None = None
        try:
            while True:
                message = self.queue.get_nowait()
                if message[0] == "log":
                    chunks.append(message[1])
                elif message[0] == "done":
                    finished = message
        except queue.Empty:
            pass

        if chunks and self.active_page is not None:
            self.active_page.append_log("".join(chunks))
        if finished is not None:
            self._finish(*finished[1:])

        self._drain_id = self.root.after(80, self._drain)

    def _finish(self, code: int, kind: str, seconds: float) -> None:
        self.busy = False
        page = self.active_page
        if page is not None:
            page.set_running(False)
        for button in self.nav_buttons:
            button.configure(state="normal")  # state=disabled 时 bg 不生效，先恢复再上色
        self._highlight(getattr(self, "_current", 0))
        self.active_page = None

        elapsed = f"{seconds:.1f} 秒"
        title = page.title if page is not None else "任务"
        if kind == "ok" and code == 0:
            self.set_status(f"「{title}」完成，用时 {elapsed}。", OK_GREEN)
        elif kind == "cancel":
            self.set_status(f"「{title}」已取消，用时 {elapsed}。")
        elif kind == "error":
            self.set_status(
                f"「{title}」出现未预料的异常，用时 {elapsed}。请点「复制日志」发给小埃。",
                WARN_RED,
            )
        else:
            self.set_status(f"「{title}」执行失败（退出码 {code}），用时 {elapsed}。", WARN_RED)

    def set_status(self, text: str, colour: str = TEXT) -> None:
        self.status.configure(text=text, fg=colour)

    # -- 关闭 -----------------------------------------------------------

    def dispose(self) -> None:
        """取消定时刷新并销毁本 App 建的全部控件。

        :meth:`on_close` 会直接摧毁整个 Tk 根窗口，但那种做法没法用于
        「同一个根窗口上反复建 App」的场景（单元测试就是），所以单独留一个
        只清自己东西的出口。
        """
        if self._drain_id is not None:
            try:
                self.root.after_cancel(self._drain_id)
            except Exception:  # pragma: no cover - 根窗口已销毁
                pass
            self._drain_id = None
        frame = getattr(self, "root_frame", None)
        if frame is not None:
            try:
                frame.destroy()
            except Exception:  # pragma: no cover
                pass
        self.pages = []

    def on_close(self) -> None:
        if self.busy:
            keep = messagebox.askyesno(
                "任务还在跑",
                "后台还有任务没结束，现在退出可能留下不完整的输出文件。\n确定要退出吗？",
                parent=self.root,
                default="no",
            )
            if not keep:
                return
        self.dispose()
        self.root.destroy()


# ------------------------------------------------------------------ 入口


def launch() -> int:
    """打开图形界面并阻塞到窗口关闭。

    :raises GuiUnavailable: 没有 tkinter，或创建窗口失败（例如无图形界面）
    """
    if not dialog_available():
        raise GuiUnavailable(
            "当前 Python 环境里没有 tkinter，开不了图形界面。\n"
            "请在项目目录执行 uv sync 重装依赖后再试；"
            "也可以先用命令行菜单：uv run speedup menu"
        )

    _enable_dpi_awareness()  # 必须在 Tk() 之前
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        raise GuiUnavailable(f"创建窗口失败（可能是当前环境没有图形界面）：{exc}") from exc

    factor = _apply_scaling(root)
    app = SpeedupApp(root, factor)
    try:
        root.mainloop()
    finally:
        try:
            root.destroy()
        except Exception:  # pragma: no cover - 已经销毁过了
            pass
        del app
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    """``speedup gui`` 的入口：强制开图形界面，开不了就明确报错。"""
    try:
        return launch()
    except GuiUnavailable as exc:
        console.fail(str(exc))
        return 2


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
