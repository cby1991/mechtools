"""系统交互工具测试。

真去开资源管理器会把测试机搞乱，所以这里把 ``os.startfile`` /
``subprocess.Popen`` 打桩，只验证「选对了分支、失败也不炸」。
"""

from __future__ import annotations

import subprocess
import sys

import pytest

from speedup import osutil


class TestIsWindows:
    def test_returns_bool(self):
        assert isinstance(osutil.is_windows(), bool)

    def test_matches_platform(self):
        assert osutil.is_windows() == sys.platform.startswith("win")


class TestOpenWithDefaultApp:
    def test_uses_startfile_on_windows(self, monkeypatch, tmp_path):
        calls = []
        monkeypatch.setattr(osutil, "is_windows", lambda: True)
        monkeypatch.setattr("os.startfile", lambda path: calls.append(path), raising=False)

        assert osutil.open_with_default_app(tmp_path / "a.png") is True
        assert len(calls) == 1

    def test_returns_false_when_startfile_fails(self, monkeypatch, tmp_path):
        def boom(path):
            raise OSError("没有关联程序")

        monkeypatch.setattr(osutil, "is_windows", lambda: True)
        monkeypatch.setattr("os.startfile", boom, raising=False)

        assert osutil.open_with_default_app(tmp_path / "a.weird") is False

    @pytest.mark.parametrize(
        ("platform", "opener"),
        [("darwin", "open"), ("linux", "xdg-open")],
    )
    def test_uses_platform_opener_on_posix(self, monkeypatch, tmp_path, platform, opener):
        commands = []
        monkeypatch.setattr(osutil, "is_windows", lambda: False)
        monkeypatch.setattr(osutil.sys, "platform", platform)
        monkeypatch.setattr(
            osutil.subprocess, "Popen", lambda cmd, **kwargs: commands.append(cmd)
        )

        assert osutil.open_with_default_app(tmp_path / "a.png") is True
        assert commands[0][0] == opener

    def test_returns_false_when_opener_missing(self, monkeypatch, tmp_path):
        def boom(cmd, **kwargs):
            raise OSError("xdg-open 不存在")

        monkeypatch.setattr(osutil, "is_windows", lambda: False)
        monkeypatch.setattr(osutil.subprocess, "Popen", boom)

        assert osutil.open_with_default_app(tmp_path / "a.png") is False

    def test_accepts_pathlib_paths(self, monkeypatch, tmp_path):
        calls = []
        monkeypatch.setattr(osutil, "is_windows", lambda: True)
        monkeypatch.setattr("os.startfile", lambda path: calls.append(path), raising=False)

        osutil.open_with_default_app(tmp_path)
        assert isinstance(calls[0], str)


class TestRevealInFileManager:
    def test_windows_opens_the_folder(self, monkeypatch, tmp_path):
        calls = []
        monkeypatch.setattr(osutil, "is_windows", lambda: True)
        monkeypatch.setattr("os.startfile", lambda path: calls.append(path), raising=False)

        assert osutil.reveal_in_file_manager(tmp_path / "a.png") is True
        assert calls == [str(tmp_path / "a.png")]

    def test_posix_delegates_to_parent_folder(self, monkeypatch, tmp_path):
        opened = []
        monkeypatch.setattr(osutil, "is_windows", lambda: False)
        monkeypatch.setattr(osutil, "open_with_default_app", lambda p: opened.append(p) or True)

        target = tmp_path / "sub" / "a.png"
        assert osutil.reveal_in_file_manager(target) is True
        assert opened == [str(tmp_path / "sub")]


class TestModuleEntryPoint:
    def test_python_m_speedup_reports_version(self):
        """``python -m speedup`` 必须能跑（覆盖 __main__.py）。"""
        result = subprocess.run(
            [sys.executable, "-m", "speedup", "--version"],
            capture_output=True,
            text=True,
            timeout=60,
        )
        assert result.returncode == 0
        assert "speedup" in result.stdout

    def test_python_m_speedup_help(self):
        result = subprocess.run(
            [sys.executable, "-m", "speedup", "--help"],
            capture_output=True,
            text=True,
            timeout=60,
        )
        assert result.returncode == 0
        assert "车间提速工具集" in result.stdout
