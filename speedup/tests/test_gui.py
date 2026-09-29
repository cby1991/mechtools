"""图形界面测试。

这里**不启动 mainloop**（会挂住）。做法是：有一个被 ``withdraw()`` 隐藏起来的
真实 Tk 根窗口，在上面把 :class:`speedup.gui.SpeedupApp` 建起来，然后直接
调用它的方法。这样测的是真实控件、真实取值路径，而不是模拟对象。

后台任务那条链路（子线程 + 队列 + 主线程搬运）用 ``_pump`` 手动推进队列来测，
覆盖了 :meth:`SpeedupApp._worker` 的每一个分支。
"""

from __future__ import annotations

import queue
import sys
import time

import pytest

from speedup import cli, gui
from speedup.errors import SpeedupError, UserCancelled
from speedup.tools import TOOLS

# ------------------------------------------------------------------ 夹具


@pytest.fixture(scope="module")
def tk_root():
    """一个隐藏的 Tk 根窗口；整个测试模块共用一个（多开会互相干扰）。"""
    tk = pytest.importorskip("tkinter")
    try:
        root = tk.Tk()
    except tk.TclError as exc:  # 没有图形界面（纯 CI 容器等）
        pytest.skip(f"当前环境没有可用的图形界面：{exc}")
    root.withdraw()  # 别弹到用户脸上
    yield root
    try:
        root.destroy()
    except Exception:  # pragma: no cover
        pass


@pytest.fixture
def app(tk_root):
    instance = gui.SpeedupApp(tk_root)
    yield instance
    instance.dispose()


def tool_page(instance: gui.SpeedupApp, name: str) -> gui.ToolPage:
    for page in instance.pages:
        if isinstance(page, gui.ToolPage) and page.module.NAME == name:
            return page
    raise AssertionError(f"图形界面里没有 {name} 这一页")


def _pump(instance: gui.SpeedupApp, timeout: float = 10.0) -> None:
    """手动推进「队列 → 日志框」的搬运，直到后台任务收尾。"""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        instance._drain()
        if not instance.busy:
            return
        time.sleep(0.02)
    raise AssertionError("后台任务没有在预期时间内结束")


def _log_of(page: gui.BasePage) -> str:
    assert page.log is not None
    return page.log.get("1.0", "end-1c")


# ------------------------------------------------------------------ 表单与 argparse 的一致性


class TestFieldContract:
    """界面表单和 argparse 定义必须始终对得上，否则界面会悄悄改行为。"""

    def test_every_tool_has_gui_fields(self):
        missing = sorted(set(TOOLS) - set(gui.TOOL_FIELDS))
        assert not missing, f"这些工具没配图形界面表单，界面上会看不到：{missing}"

    def test_no_stale_gui_fields(self):
        stale = sorted(set(gui.TOOL_FIELDS) - set(TOOLS))
        assert not stale, f"这些表单指向了不存在的工具：{stale}"

    def test_every_dest_exists_in_argparse(self):
        parser = cli.build_parser()
        for name, fields in gui.TOOL_FIELDS.items():
            defaults = vars(parser.parse_args([name]))
            for field in fields:
                assert field.dest in defaults, f"{name} 的字段 {field.dest!r} 不是 argparse 的属性"

    def test_every_flag_is_accepted(self):
        """逐个开关喂给真解析器，写错了会 SystemExit。"""
        parser = cli.build_parser()
        for name, fields in gui.TOOL_FIELDS.items():
            for field in fields:
                if field.flag is None:
                    continue
                if field.kind == "bool":
                    argv = [name, field.flag]
                elif field.kind == "choice":
                    value = (field.choices or ((None, field.default),))[0][1]
                    argv = [name, field.flag, str(value)]
                else:
                    argv = [name, field.flag, "x"]
                parser.parse_args(argv)  # 不认识就会 SystemExit

    def test_positional_fields_come_first_in_declaration_order(self):
        """位置参数的先后顺序必须和 argparse 里 add_argument 的顺序一致。"""
        parser = cli.build_parser()
        for name, fields in gui.TOOL_FIELDS.items():
            positionals = [f.dest for f in fields if f.flag is None and f.kind != "bool"]
            if not positionals:
                continue
            args = vars(parser.parse_args([name]))
            # 逐个喂值，确认第 n 个位置参数确实落在第 n 个 dest 上
            for index, dest in enumerate(positionals):
                argv = [name, *[f"第{i}个" for i in range(index + 1)]]
                assert vars(parser.parse_args(argv))[dest].startswith(f"第{index}个"), (
                    f"{name} 的位置参数顺序对不上：第 {index + 1} 个应为 {dest}"
                )


# ------------------------------------------------------------------ argv 拼装


class TestBuildArgv:
    def test_auto_open_is_on_by_default(self, app):
        """勾选框默认勾上 = 不传 ``--no-open`` = 生成后自动打开。"""
        page = tool_page(app, "gbbuild")
        page.set_value("folder", r"D:\某文件夹")
        assert page.value_of("no_open") is True
        assert page.build_argv() == [r"D:\某文件夹"]

    def test_unchecking_auto_open_adds_the_negated_flag(self, app):
        page = tool_page(app, "gbbuild")
        page.set_value("folder", r"D:\某文件夹")
        page.set_value("no_open", False)
        assert page.build_argv() == [r"D:\某文件夹", "--no-open"]

    def test_optional_text_becomes_a_flag(self, app):
        page = tool_page(app, "gbbuild")
        page.set_value("folder", "D:/x")
        page.set_value("out", "清单.xlsx")
        assert page.build_argv() == ["D:/x", "--out", "清单.xlsx"]

    def test_gbcopy_positional_order(self, app):
        """excel 必须排在 search_dir 前面，否则文件会被拷到错误的地方。"""
        page = tool_page(app, "gbcopy")
        page.set_value("excel", r"D:\a\清单.xlsx")
        page.set_value("search_dir", r"D:\b")
        assert page.build_argv() == [
            r"D:\a\清单.xlsx",
            r"D:\b",
            "--min-score",
            "3",
        ]

    def test_gbcopy_dry_run_and_loose_match(self, app):
        page = tool_page(app, "gbcopy")
        page.set_value("excel", "a.xlsx")
        page.set_value("search_dir", "b")
        page.set_value("dry_run", True)
        page.set_value("min_score", 1)
        argv = page.build_argv()
        assert "--dry-run" in argv
        assert argv[argv.index("--min-score") + 1] == "1"

    def test_choice_accepts_display_text_too(self, app):
        page = tool_page(app, "gbcopy")
        page.set_value("min_score", "完全相同")
        assert page.value_of("min_score") == 4

    def test_unknown_choice_is_rejected(self, app):
        page = tool_page(app, "gbcopy")
        with pytest.raises(ValueError, match="没有这个选项"):
            page.set_value("min_score", 2)

    def test_missing_required_field_blocks(self, app):
        page = tool_page(app, "gbbuild")
        page.set_value("folder", "   ")
        assert page.can_run() is not None
        with pytest.raises(ValueError, match="目标文件夹"):
            page.build_argv()

    def test_whitespace_is_trimmed(self, app):
        page = tool_page(app, "gbbuild")
        page.set_value("folder", "  D:/x  ")
        assert page.build_argv()[0] == "D:/x"

    def test_reset_clears_inputs(self, app):
        page = tool_page(app, "gbbuild")
        page.set_value("folder", "D:/x")
        page.set_value("no_open", False)
        page.reset()
        assert page.value_of("folder") == ""
        assert page.value_of("no_open") is True
        assert page.can_run() is not None

    def test_unknown_dest_raises(self, app):
        with pytest.raises(KeyError):
            tool_page(app, "gbbuild").field("没有这个字段")

    @pytest.mark.parametrize("name", sorted(TOOLS))
    def test_argv_survives_the_real_parser(self, app, name):
        """最重要的一条：界面拼出来的 argv 必须能被真正的 argparse 吃下去。"""
        page = tool_page(app, name)
        for field in page.fields:
            if field.required or field.kind == "folder":
                page.set_value(field.dest, "D:/some/folder" if field.kind == "folder" else "x")
        argv = page.build_argv()
        args = cli.build_parser().parse_args([name, *argv])
        assert args.command == name


# ------------------------------------------------------------------ 页面结构


class TestPages:
    def test_pages_cover_tools_and_extras(self, app):
        titles = [page.title for page in app.pages]
        for module in TOOLS.values():
            assert module.TITLE in titles
        assert "环境体检" in titles
        assert "历史脚本归档" in titles
        assert "使用说明" in titles

    def test_only_one_page_visible_at_a_time(self, app):
        for index in range(len(app.pages)):
            app.show(index)
            visible = [page for page in app.pages if page.winfo_manager() == "pack"]
            assert len(visible) == 1
            assert visible[0] is app.pages[index]

    def test_help_page_has_no_log_widget(self, app):
        help_page = next(page for page in app.pages if page.title == "使用说明")
        assert help_page.log is None
        help_page.append_log("不会炸")  # 得能安全地空转
        help_page.clear_log()

    def test_help_page_explains_every_tool(self, app):
        help_page = next(page for page in app.pages if page.title == "使用说明")
        assert help_page.text is not None
        text = help_page.text.get("1.0", "end")
        for module in TOOLS.values():
            assert module.TITLE in text, f"使用说明里没提 {module.TITLE}"

    def test_every_tool_page_has_a_run_button(self, app):
        for page in app.pages:
            if isinstance(page, gui.ToolPage):
                assert str(page.run_button.cget("state")) == "normal"

    def test_action_pages_are_wired(self, app):
        doctor = next(page for page in app.pages if page.title == "环境体检")
        archive = next(page for page in app.pages if page.title == "历史脚本归档")
        assert callable(doctor.run_job())
        assert callable(archive.run_job())


# ------------------------------------------------------------------ 输出搬运


class TestQueueWriter:
    def test_writes_land_in_the_queue(self):
        sink: queue.Queue[tuple] = queue.Queue()
        writer = gui._QueueWriter(sink)
        assert writer.write("你好") == 2
        assert sink.get_nowait() == ("log", "你好")

    def test_empty_write_is_ignored(self):
        sink: queue.Queue[tuple] = queue.Queue()
        gui._QueueWriter(sink).write("")
        assert sink.empty()

    def test_encoding_is_utf8_so_doctor_does_not_warn(self):
        """日志直接进 Tk 控件，不做编码转换；报 utf-8 才不会触发体检的假警告。"""
        assert gui._QueueWriter(queue.Queue()).encoding == "utf-8"


class TestTaskRunner:
    def test_success_streams_output_and_reports_done(self, app):
        page = tool_page(app, "gbbuild")

        def job() -> int:
            print("hello from tool")
            print("第二行")
            return 0

        app.run_task(page, job)
        _pump(app)
        text = _log_of(page)
        assert "hello from tool" in text
        assert "第二行" in text
        assert app.busy is False
        assert "完成" in app.status.cget("text")
        assert str(page.run_button.cget("state")) == "normal"

    def test_starting_a_task_locks_the_ui(self, app):
        page = tool_page(app, "gbbuild")
        gate = {"release": False}

        def job() -> int:
            while not gate["release"]:
                time.sleep(0.01)
            return 0

        app.run_task(page, job)
        try:
            assert app.busy is True
            assert str(page.run_button.cget("state")) == "disabled"
            assert all(str(b.cget("state")) == "disabled" for b in app.nav_buttons)
            app.show(0)  # 跑任务时不许切页
            assert "等它结束" in app.status.cget("text")
        finally:
            gate["release"] = True
        _pump(app)

    def test_speedup_error_is_reported(self, app):
        page = tool_page(app, "gbbuild")

        def job() -> int:
            raise SpeedupError("磁盘满了")

        app.run_task(page, job)
        _pump(app)
        assert "[失败] 磁盘满了" in _log_of(page)
        assert "失败" in app.status.cget("text")

    def test_cancellation_is_reported(self, app):
        page = tool_page(app, "gbbuild")

        def job() -> int:
            raise UserCancelled("用户未选择文件夹")

        app.run_task(page, job)
        _pump(app)
        assert "[取消]" in _log_of(page)
        assert "已取消" in app.status.cget("text")

    def test_unexpected_exception_is_dumped_with_traceback(self, app):
        page = tool_page(app, "gbbuild")

        def job() -> int:
            raise RuntimeError("真 bug")

        app.run_task(page, job)
        _pump(app)
        text = _log_of(page)
        assert "[异常]" in text
        assert "真 bug" in text
        assert "Traceback" in text  # 有堆栈才好定位
        assert "复制日志" in app.status.cget("text")

    def test_argparse_systemexit_is_absorbed(self, app):
        """参数解析失败会抛 SystemExit（BaseException），不能被漏过去。"""
        page = tool_page(app, "gbbuild")

        def job() -> int:
            raise SystemExit(2)

        app.run_task(page, job)
        _pump(app)
        assert "[退出]" in _log_of(page)
        assert "退出码 2" in app.status.cget("text")

    def test_non_zero_exit_code_is_reported(self, app):
        page = tool_page(app, "gbbuild")
        app.run_task(page, lambda: 3)
        _pump(app)
        assert "退出码 3" in app.status.cget("text")

    def test_log_can_be_cleared_and_copied(self, app):
        page = tool_page(app, "gbbuild")
        app.run_task(page, lambda: print("一些输出") or 0)
        _pump(app)
        assert _log_of(page).strip()
        page.copy_log()
        assert page.app.root.clipboard_get() == _log_of(page)
        page.clear_log()
        assert not _log_of(page).strip()

    def test_copy_empty_log_is_a_noop(self, app):
        page = tool_page(app, "gbbuild")
        page.clear_log()
        page.copy_log()  # 不该炸
        assert app.status.cget("text") != "日志已复制到剪贴板。"


class TestNavigation:
    def test_hover_does_not_override_the_active_highlight(self, app):
        app.show(0)
        active_bg = app.nav_buttons[0].cget("bg")
        app._hover(0, True)
        assert app.nav_buttons[0].cget("bg") == active_bg

    def test_hover_tints_inactive_buttons(self, app):
        app.show(0)
        app._hover(1, True)
        assert app.nav_buttons[1].cget("bg") == gui.SOFT_HOVER
        app._hover(1, False)
        assert app.nav_buttons[1].cget("bg") == gui.SOFT


class TestFonts:
    """日志区的字体不是等宽的话，``console`` 渲染的表格会整片错位。"""

    def test_log_font_is_monospace(self, app):
        from tkinter import font as tkfont

        log_font = tkfont.Font(root=app.root, font=app.font_mono)
        assert log_font.measure("W") == log_font.measure("i"), (
            f"日志字体 {log_font.actual()['family']!r} 不是等宽，表格会错位"
        )

    def test_cjk_is_exactly_twice_as_wide(self, app):
        """``console.display_width`` 把中日韩文字算 2 列，字体必须真的满足这个前提。"""
        from tkinter import font as tkfont

        log_font = tkfont.Font(root=app.root, font=app.font_mono)
        assert log_font.measure("中") == log_font.measure("W") * 2

    def test_ui_font_is_resolved_to_something_real(self, app):
        from tkinter import font as tkfont

        for spec in (app.font_body, app.font_bold, app.font_title, app.font_small):
            actual = tkfont.Font(root=app.root, font=spec).actual()
            assert actual["family"], f"{spec} 没解析出字体家族"
            assert int(actual["size"]) > 0


# ------------------------------------------------------------------ pythonw 模式


class TestPythonwMode:
    """双击快捷方式走的是 ``pythonw.exe``，那里 ``sys.stdout`` 是 ``None``。"""

    def test_ensure_streams_fills_none_stdout(self, monkeypatch):
        monkeypatch.setattr(sys, "stdout", None)
        monkeypatch.setattr(sys, "stderr", None)
        cli.console.ensure_streams()
        assert sys.stdout is not None
        assert sys.stderr is sys.stdout
        print("这句话不应该炸")  # 关键的回归点

    def test_ensure_streams_keeps_existing_streams(self):
        original = sys.stdout
        cli.console.ensure_streams()
        assert sys.stdout is original


# ------------------------------------------------------------------ 入口路由


class TestMainRouting:
    def test_gui_subcommand_runs_the_window(self, monkeypatch):
        monkeypatch.setattr(cli, "_try_gui", lambda: (0, ""))
        assert cli.main(["gui"]) == 0

    def test_gui_subcommand_reports_failure_loudly(self, monkeypatch, capsys):
        monkeypatch.setattr(cli, "_try_gui", lambda: (None, "没有 tkinter"))
        assert cli.main(["gui"]) == 2
        assert "没有 tkinter" in capsys.readouterr().out

    def test_bare_invocation_opens_the_gui(self, monkeypatch):
        monkeypatch.setattr(cli, "_try_gui", lambda: (0, ""))
        assert cli.main([]) == 0

    def test_bare_invocation_falls_back_to_menu(self, monkeypatch, capsys):
        monkeypatch.setattr(cli, "_try_gui", lambda: (None, "开不了窗口"))
        monkeypatch.setattr(cli, "run_menu", lambda: 7)
        assert cli.main([]) == 7
        assert "开不了窗口" in capsys.readouterr().out

    def test_menu_subcommand_never_opens_a_window(self, monkeypatch):
        calls: list[int] = []
        monkeypatch.setattr(cli, "_try_gui", lambda: calls.append(1) or (0, ""))
        monkeypatch.setattr(cli, "run_menu", lambda: 0)
        assert cli.main(["menu"]) == 0
        assert calls == [], "menu 子命令不该去开图形界面"

    def test_no_console_falls_back_to_an_alert_box(self, monkeypatch):
        monkeypatch.setattr(cli, "_try_gui", lambda: (None, "开不了窗口"))
        monkeypatch.setattr(sys, "stdout", None)
        seen: dict[str, str] = {}
        monkeypatch.setattr(cli, "_alert_box", lambda message: seen.setdefault("msg", message))
        assert cli.main([]) == 2
        assert "开不了窗口" in seen["msg"]

    def test_try_gui_reports_a_friendly_reason_when_the_module_breaks(self, monkeypatch):
        import builtins

        real_import = builtins.__import__

        def fake_import(name, *args, **kwargs):
            if name == "gui":
                raise ImportError("no module named tkinter")
            return real_import(name, *args, **kwargs)

        monkeypatch.setattr(builtins, "__import__", fake_import)
        code, reason = cli._try_gui()
        assert code is None
        assert "图形界面模块加载失败" in reason
        assert "tkinter" in reason
