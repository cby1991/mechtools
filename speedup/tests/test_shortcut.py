"""快捷方式测试。

分两层：
* **纯计算部分**（该指向什么、路径怎么拼）任何环境都能测；
* **真的写一个 .lnk** 只在 Windows 上测，而且要真调 COM —— 这才是这段代码
  最容易出错的地方（虚表下标写错、指针类型推断错都会当场炸）。
"""

from __future__ import annotations

import os

import pytest

from speedup import cli, shortcut
from speedup.config import PROJECT_ROOT
from speedup.errors import SpeedupError
from speedup.osutil import desktop_dir, is_windows

windows_only = pytest.mark.skipif(not is_windows(), reason="快捷方式只在 Windows 上有意义")


def lnk_strings(raw: bytes) -> str:
    """把 ``.lnk`` 里可能藏字符串的地方都翻出来。

    路径有两种存法，两种都得看：

    * **StringData 段**（UTF-16LE）—— 工作目录、参数、图标、名称；
    * **LinkInfo 段**（ANSI）—— 绝对目标路径。目标与快捷方式**跨盘符**时写不出
      相对路径，``HasRelativePath`` 标志位会关掉，此时目标路径只存在于 LinkInfo 里。

    用 latin-1 而不是 mbcs：它是无损的字节到字符映射，ASCII 内容照样能找到，
    又不依赖系统代码页。
    """
    return raw.decode("utf-16-le", errors="ignore") + raw.decode("latin-1", errors="ignore")


class TestPlan:
    def test_targets_pythonw_not_python(self):
        """必须指向 pythonw.exe —— 用 python.exe 会弹出一个黑框控制台。"""
        spec = shortcut.plan_shortcut()
        assert os.path.basename(spec.target).lower() == "pythonw.exe"

    def test_arguments_run_the_gui_entry_point(self):
        assert shortcut.plan_shortcut().arguments == "-m speedup"

    def test_working_dir_is_the_project_root(self):
        assert shortcut.plan_shortcut().working_dir == str(PROJECT_ROOT)

    def test_icon_exists_and_is_the_packaged_one(self):
        """图标没生成的话，快捷方式会顶着 Python 默认图标，所以这里盯死。"""
        plan = shortcut.plan_shortcut()
        assert plan.icon is not None, "没找到图标，先跑 examples/make_icon.py"
        assert plan.icon == str(shortcut.icon_path())
        assert os.path.isfile(plan.icon)

    def test_filename_gets_lnk_suffix(self):
        assert shortcut.plan_shortcut().filename == f"{shortcut.SHORTCUT_NAME}.lnk"

    def test_custom_name(self):
        assert shortcut.plan_shortcut("小埃").filename == "小埃.lnk"

    def test_describe_mentions_every_field(self):
        text = shortcut.describe(shortcut.plan_shortcut())
        for piece in ("pythonw.exe", "-m speedup", str(PROJECT_ROOT), "speedup.ico"):
            assert piece in text


class TestDestination:
    def test_default_is_the_project_root(self):
        assert shortcut.default_destination() == PROJECT_ROOT

    def test_shortcut_path_joins_the_filename(self, tmp_path):
        spec = shortcut.plan_shortcut()
        assert shortcut.shortcut_path(spec, tmp_path) == tmp_path / spec.filename

    def test_missing_directory_is_rejected(self, tmp_path):
        spec = shortcut.plan_shortcut()
        with pytest.raises(SpeedupError, match="保存目录不存在"):
            shortcut.shortcut_path(spec, tmp_path / "没有这个目录")

    @windows_only
    def test_desktop_exists_or_is_reported(self):
        found = shortcut.default_destination(desktop=True)
        assert found.is_dir()

    @windows_only
    def test_desktop_dir_is_expanded(self):
        """装了 OneDrive 的机器上桌面会被重定向，值里可能带 %USERPROFILE%。"""
        found = desktop_dir()
        assert found is not None and os.path.isdir(found)
        assert "%" not in found


@windows_only
class TestWriteLnk:
    def test_creates_a_real_shortcut(self, tmp_path):
        spec = shortcut.plan_shortcut()
        path = shortcut.create_shortcut(spec, tmp_path / spec.filename)

        assert path.is_file()
        raw = path.read_bytes()
        # 0x4C 是 .lnk 的固定头
        assert int.from_bytes(raw[:4], "little") == 0x4C

        text = lnk_strings(raw)
        assert "pythonw.exe" in text
        assert "-m speedup" in text
        assert str(PROJECT_ROOT) in text
        assert "speedup.ico" in text

    def test_cross_drive_target_still_resolves(self, tmp_path):
        """临时目录在 C:、目标在 D: —— 这时没有相对路径可写，得靠 LinkInfo 兜底。

        这是一条回归测试：曾经以为「字符串搜不到」是写坏了，其实是跨盘符的正常表现。
        """
        spec = shortcut.plan_shortcut()
        raw = shortcut.create_shortcut(spec, tmp_path / spec.filename).read_bytes()
        flags = int.from_bytes(raw[20:24], "little")
        has_relative_path = bool(flags & 0x08)
        assert has_relative_path == (
            os.path.splitdrive(spec.target)[0].lower()
            == os.path.splitdrive(str(tmp_path))[0].lower()
        )
        assert spec.target in lnk_strings(raw)  # 无论如何绝对路径都要在

    def test_overwriting_an_existing_shortcut_is_fine(self, tmp_path):
        spec = shortcut.plan_shortcut()
        first = shortcut.create_shortcut(spec, tmp_path / spec.filename)
        size = first.stat().st_size
        second = shortcut.create_shortcut(spec, tmp_path / spec.filename)
        assert second == first
        assert second.stat().st_size == size

    def test_creates_missing_parent_directories(self, tmp_path):
        spec = shortcut.plan_shortcut()
        nested = tmp_path / "a" / "b" / spec.filename
        assert shortcut.create_shortcut(spec, nested).is_file()

    def test_no_icon_still_works(self, tmp_path):
        """没图标也得能建出来，不能因为缺素材就整个失败。"""
        base = shortcut.plan_shortcut()
        plain = shortcut.ShortcutSpec(
            name=base.name,
            target=base.target,
            arguments=base.arguments,
            working_dir=base.working_dir,
            icon=None,
            description=base.description,
        )
        path = shortcut.create_shortcut(plain, tmp_path / plain.filename)
        assert path.is_file()


class TestCliIntegration:
    def test_subcommand_parses(self):
        args = cli.build_parser().parse_args(["shortcut", "--desktop"])
        assert args.command == "shortcut"
        assert args.desktop is True
        assert args.dest is None

    @windows_only
    def test_main_creates_the_shortcut_and_returns_zero(self, tmp_path, capsys):
        code = cli.main(["shortcut", "--dest", str(tmp_path)])
        assert code == 0
        assert (tmp_path / f"{shortcut.SHORTCUT_NAME}.lnk").is_file()
        out = capsys.readouterr().out
        assert "已创建" in out
        assert "pythonw.exe" in out  # 动手前先把内容打给人核对

    @windows_only
    def test_main_reports_a_bad_destination(self, tmp_path, capsys):
        code = cli.main(["shortcut", "--dest", str(tmp_path / "不存在")])
        assert code == 2
        assert "[失败]" in capsys.readouterr().out

    @windows_only
    def test_main_reports_overwriting(self, tmp_path, capsys):
        cli.main(["shortcut", "--dest", str(tmp_path)])
        capsys.readouterr()
        cli.main(["shortcut", "--dest", str(tmp_path)])
        assert "已覆盖" in capsys.readouterr().out
