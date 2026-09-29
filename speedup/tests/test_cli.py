"""CLI 测试：参数解析、退出码、交互菜单。

菜单不是另写一套逻辑，所以这里既测「输入编号能不能选中工具」，
也测「选中的工具收到的参数和命令行一致」。
"""

from __future__ import annotations

import os

import pytest

from speedup import cli, console
from speedup.tools import TOOLS


class TestParser:
    def test_gbbuild_defaults(self):
        args = cli.build_parser().parse_args(["gbbuild"])
        assert args.folder is None
        assert args.out is None
        assert args.no_open is False

    def test_gbbuild_with_path_and_flags(self):
        args = cli.build_parser().parse_args(["gbbuild", r"D:\x", "--out", "a.xlsx", "--no-open"])
        assert args.folder == r"D:\x"
        assert args.out == "a.xlsx"
        assert args.no_open is True

    def test_gbcopy_defaults(self):
        args = cli.build_parser().parse_args(["gbcopy"])
        assert args.excel is None
        assert args.search_dir is None
        assert args.dry_run is False
        assert args.min_score == 3

    def test_gbcopy_min_score_choices(self):
        parser = cli.build_parser()
        for value in (1, 3, 4):
            assert parser.parse_args(["gbcopy", "--min-score", str(value)]).min_score == value
        with pytest.raises(SystemExit):
            parser.parse_args(["gbcopy", "--min-score", "2"])

    def test_unknown_command_exits(self):
        with pytest.raises(SystemExit):
            cli.build_parser().parse_args(["没有这个命令"])

    def test_every_tool_is_registered(self):
        parser = cli.build_parser()
        for name in TOOLS:
            assert parser.parse_args([name]).command == name

    def test_tools_are_wired_to_the_same_run_path(self, monkeypatch):
        """回归测试：菜单和命令行必须走 module.run，不能被 argparse 提前绑死。"""
        called = {}

        def fake_run(args):
            called["ok"] = True
            return 0

        monkeypatch.setattr("speedup.tools.gbbuild.run", fake_run)
        assert cli.main(["gbbuild", "--no-open"]) == 0
        assert called.get("ok") is True

    def test_help_mentions_uv_run(self, capsys):
        with pytest.raises(SystemExit):
            cli.build_parser().parse_args(["--help"])
        assert "uv run speedup" in capsys.readouterr().out


class TestVersion:
    def test_version_exits_zero(self, capsys):
        with pytest.raises(SystemExit) as excinfo:
            cli.main(["--version"])
        assert excinfo.value.code == 0
        assert "speedup" in capsys.readouterr().out


class TestResolution:
    def test_resolve_by_number(self):
        ordered = sorted(TOOLS.values(), key=lambda module: module.NAME)
        assert cli._resolve_choice("1") is ordered[0]
        assert cli._resolve_choice(str(len(ordered))) is ordered[-1]

    def test_resolve_by_name(self):
        assert cli._resolve_choice("gbbuild") is TOOLS["gbbuild"]

    def test_out_of_range(self):
        assert cli._resolve_choice("99") is None
        assert cli._resolve_choice("0") is None

    def test_unknown_name(self):
        assert cli._resolve_choice("nope") is None


class TestDoctor:
    def test_runs_and_reports(self, capsys):
        code = cli.run_doctor()
        out = capsys.readouterr().out
        assert code in (0, 1)
        assert "环境体检" in out
        assert "tkinter" in out

    def test_returns_one_when_something_missing(self, monkeypatch, capsys):
        monkeypatch.setattr(cli, "_probe_tkinter", lambda: "不可用 —— 没有图形界面")
        # 让体检真的用到被替换的探针
        real_probe = cli._probe

        def fake_probe(name, probe):
            if name == "tkinter 弹窗":
                return name, "不可用 —— 没有图形界面"
            return real_probe(name, probe)

        monkeypatch.setattr(cli, "_probe", fake_probe)
        assert cli.run_doctor() == 1
        assert "不可用" in capsys.readouterr().out

    def test_reports_pillow(self, capsys):
        """pillow 是嵌轴测图必需的 —— 缺了它图片会**静默不嵌**，最难查。

        所以它必须在体检清单里，让人一眼看出问题在哪。
        """
        cli.run_doctor()
        assert "pillow" in capsys.readouterr().out

    def test_reports_optional_cadquery(self, capsys):
        """cadquery 是**可选**依赖，但体检要把它列出来。

        不列的话，用户不知道自己的机器到底能不能出轴测图 ——
        只能靠自己试一遍才发现「原来没装」。
        """
        cli.run_doctor()
        assert "cadquery" in capsys.readouterr().out

    def test_missing_optional_dependency_is_not_a_failure(self, monkeypatch, capsys):
        """可选依赖缺失**不该**让体检返回 1。

        它只影响出图这一个功能，装不装都能正常出清单 ——
        判成「失败」会让人以为整个环境坏了（见 R29）。
        """
        real_probe = cli._probe

        def fake_probe(name, probe):
            if "cadquery" in name:
                return name, "未安装 —— 轴测图不可用（可选）"
            return real_probe(name, probe)

        monkeypatch.setattr(cli, "_probe", fake_probe)
        assert cli.run_doctor() == 0
        assert "cadquery" in capsys.readouterr().out


class TestPickInstaller:
    """找一个能装包的入口。

    本项目用 **uv** 建 venv，默认**不带 pip** —— 所以不能想当然地
    用 ``python -m pip``，得先看 venv 里到底有没有 pip。
    """

    def test_prefers_venv_pip_when_present(self, monkeypatch):
        monkeypatch.setattr(cli, "_has_module", lambda name: name == "pip")
        command = cli.pick_installer()
        assert command is not None
        assert command[1:] == ["-m", "pip"], "有 pip 就用 venv 自己的 pip"

    def test_falls_back_to_uv_when_no_pip(self, monkeypatch):
        monkeypatch.setattr(cli, "_has_module", lambda name: False)
        monkeypatch.setattr(cli.shutil, "which", lambda name: "C:/uv/uv.exe" if name == "uv" else None)
        command = cli.pick_installer()
        assert command == ["C:/uv/uv.exe", "pip"]

    def test_returns_none_when_nothing_available(self, monkeypatch):
        """两样都没有时要老实返回 None，让调用方去提示手装 —— 不能崩。"""
        monkeypatch.setattr(cli, "_has_module", lambda name: False)
        monkeypatch.setattr(cli.shutil, "which", lambda name: None)
        monkeypatch.setattr(cli, "_bundled_uv", lambda: None)
        assert cli.pick_installer() is None

    def test_uses_bundled_uv_as_last_resort(self, monkeypatch):
        """本机 uv 不在 PATH，但有个已知的绝对路径 —— 兜底用它。"""
        monkeypatch.setattr(cli, "_has_module", lambda name: False)
        monkeypatch.setattr(cli.shutil, "which", lambda name: None)
        monkeypatch.setattr(cli, "_bundled_uv", lambda: "C:/known/uv.exe")
        assert cli.pick_installer() == ["C:/known/uv.exe", "pip"]


class TestInstallCommand:
    def test_pip_prefix_goes_straight_to_install(self):
        command = cli.install_command(["python.exe", "-m", "pip"], "pillow")
        assert command == ["python.exe", "-m", "pip", "install", "pillow"]

    def test_uv_gets_python_pinned_after_install(self):
        """uv 必须钉住解释器，而且 `--python` 只能在 `install` **之后**。

        实测：`uv pip --python X install pillow` 直接报 Usage；
        正确的是 `uv pip install --python X pillow`。
        """
        command = cli.install_command(["C:/uv/uv.exe", "pip"], "pillow")
        assert command == [
            "C:/uv/uv.exe",
            "pip",
            "install",
            "--python",
            cli.sys.executable,
            "pillow",
        ]

    def test_multiple_packages(self):
        command = cli.install_command(["python.exe", "-m", "pip"], "pillow", "openpyxl")
        assert command[-2:] == ["pillow", "openpyxl"]


class TestDoctorInstall:
    def test_reports_nothing_to_do_when_all_present(self, capsys):
        """依赖齐全时不该瞎装 —— 报个「就绪」就退出，别去联网。"""
        assert cli.run_doctor(install=True) == 0
        out = capsys.readouterr().out
        assert "就绪" in out
        assert "开始安装" not in out

    def test_gives_manual_hint_when_no_installer(self, monkeypatch, capsys):
        """找不到装包入口时，要给出**能照抄的命令**，而不是干说「失败」。"""
        monkeypatch.setattr(cli, "pick_installer", lambda: None)
        monkeypatch.setattr(cli, "_probe_pillow", lambda: "不可用 —— 假装缺了")
        assert cli.run_doctor(install=True) == 1
        out = capsys.readouterr().out
        assert "pillow" in out
        assert "uv pip install" in out or "pip install" in out


class TestArchiveCommand:
    """归档区是**本地留档**，不进仓库（见 `.gitignore`）。

    所以在新克隆的机器上没有 `archive/` 目录 —— 这组测试要能优雅跳过，
    而不是红一片。
    """

    @pytest.fixture(autouse=True)
    def _needs_archive(self):
        if not cli.archive_root() or not os.path.isdir(cli.archive_root()):
            pytest.skip("本机没有 archive/ 目录（归档区不进仓库，属正常情况）")

    def test_lists_categories(self, capsys):
        assert cli.run_archive() == 0
        out = capsys.readouterr().out
        assert "历史脚本归档" in out
        assert "合计" in out

    def test_counts_files_recursively(self, capsys):
        """回归测试：font_pipeline 的文件全在子目录里，只数顶层会显示 0。"""
        cli.run_archive()
        out = capsys.readouterr().out
        line = next(row for row in out.splitlines() if row.startswith("font_pipeline"))
        assert line.split()[1] != "0", f"分类统计没递归：{line!r}"

    def test_every_category_reports_a_file_count(self, capsys):
        cli.run_archive()
        out = capsys.readouterr().out
        for category in ("diag_network", "diag_windows", "font_pipeline", "build_swauto", "web", "manufacturing"):
            line = next((row for row in out.splitlines() if row.startswith(category)), None)
            assert line is not None, f"分类 {category} 没出现在归档列表里"
            assert line.split()[1].isdigit() and int(line.split()[1]) > 0, f"{category} 文件数为 0：{line!r}"


class TestMain:
    def test_gbbuild_on_missing_folder_returns_2(self, tmp_path, capsys):
        code = cli.main(["gbbuild", str(tmp_path / "没有这个目录"), "--no-open"])
        assert code == 2
        assert "[失败]" in capsys.readouterr().out

    def test_doctor_command(self, capsys):
        assert cli.main(["doctor"]) in (0, 1)
        assert "环境体检" in capsys.readouterr().out

    def test_keyboard_interrupt_is_handled(self, monkeypatch, capsys):
        def boom(args):
            raise KeyboardInterrupt

        monkeypatch.setattr("speedup.tools.gbbuild.run", boom)
        assert cli.main(["gbbuild", "--no-open"]) == 130

    def test_unexpected_exception_propagates(self, monkeypatch):
        """未预期的异常要冒泡（便于定位 bug），不能悄悄吞掉。"""

        def boom(args):
            raise RuntimeError("真 bug")

        monkeypatch.setattr("speedup.tools.gbbuild.run", boom)
        with pytest.raises(RuntimeError, match="真 bug"):
            cli.main(["gbbuild", "--no-open"])


class TestMenu:
    def _run(self, answers):
        pending = iter(answers)
        return cli.run_menu(input_func=lambda prompt: next(pending), pause=lambda: None)

    def test_zero_exits(self, capsys):
        assert self._run(["0"]) == 0
        assert "已退出" in capsys.readouterr().out

    def test_q_exits(self):
        assert self._run(["q"]) == 0

    def test_eof_exits(self, capsys):
        def raise_eof(prompt):
            raise EOFError

        assert cli.run_menu(input_func=raise_eof, pause=lambda: None) == 0

    def test_invalid_choice_then_exit(self, capsys):
        assert self._run(["abc", "0"]) == 0
        assert "没有这个选项" in capsys.readouterr().out

    def test_doctor_branch(self, capsys):
        assert self._run(["d", "0"]) == 0
        assert "环境体检" in capsys.readouterr().out

    def test_archive_branch(self, capsys):
        assert self._run(["a", "0"]) == 0
        assert "历史脚本归档" in capsys.readouterr().out

    def test_dispatches_tool_by_number(self, monkeypatch):
        called = {}
        monkeypatch.setattr("speedup.tools.gbbuild.run", lambda args: called.setdefault("hit", args))
        assert self._run(["1", "0"]) == 0
        assert "hit" in called
        # 菜单给工具的默认参数必须和命令行一致
        assert called["hit"].command == "gbbuild"
        assert called["hit"].folder is None

    def test_cancelled_tool_is_reported(self, monkeypatch, capsys):
        def cancel(args):
            raise cli.UserCancelled("用户未选择文件夹")

        monkeypatch.setattr("speedup.tools.gbbuild.run", cancel)
        assert self._run(["1", "0"]) == 0
        assert "已取消" in capsys.readouterr().out

    def test_menu_mentions_every_tool(self, capsys):
        self._run(["0"])
        out = capsys.readouterr().out
        for module in TOOLS.values():
            assert module.TITLE in out

    def test_summary_column_is_aligned(self, capsys):
        """回归测试：中文标题宽度不一，用 `f"{s:<12}"` 补空格会让说明列错位。"""
        self._run(["0"])
        out = capsys.readouterr().out

        columns = []
        for line in out.splitlines():
            for module in TOOLS.values():
                if module.SUMMARY in line:
                    columns.append(_display_index(line, module.SUMMARY))
                    break

        assert len(columns) == len(TOOLS), "有工具的说明没被渲染出来"
        assert len(set(columns)) == 1, f"说明列没对齐，起始显示列：{sorted(set(columns))}"


def _display_index(line: str, needle: str) -> int:
    """``needle`` 在 ``line`` 里的显示列（中文算 2 列，不能用 ``str.index``）。"""
    return console.display_width(line[: line.index(needle)])
