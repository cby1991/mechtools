"""弹窗与手输回退测试。

真实弹窗没法在自动化里点，所以这里：

* 测**可用性探测**（不能抛异常）
* 用 monkeypatch 把 ``dialog_available`` 设成 False，验证「没有 GUI 也不崩，会回退到手输」
* 给 :func:`ask_path` 注入假的 picker / 输入函数，把回退流程全跑一遍

真正点窗口那一下由 ``@pytest.mark.gui`` 的手工用例覆盖（默认跳过）。
"""

from __future__ import annotations

import os

import pytest

from speedup import dialogs


class TestAvailability:
    def test_dialog_available_returns_bool(self):
        assert isinstance(dialogs.dialog_available(), bool)

    def test_pick_folder_returns_none_without_tkinter(self, monkeypatch):
        monkeypatch.setattr(dialogs, "dialog_available", lambda: False)
        assert dialogs.pick_folder("随便") is None

    def test_pick_file_returns_none_without_tkinter(self, monkeypatch):
        monkeypatch.setattr(dialogs, "dialog_available", lambda: False)
        assert dialogs.pick_file("随便") is None

    @pytest.mark.skipif(not dialogs.dialog_available(), reason="需要 tkinter")
    def test_run_dialog_swallows_exceptions(self, monkeypatch):
        """哪怕 Tk 初始化直接炸，也只能返回 None，不能把工具带崩。"""

        def boom():
            raise RuntimeError("模拟 Tk 初始化失败")

        monkeypatch.setattr("tkinter.Tk", boom)
        monkeypatch.setattr(dialogs, "dialog_available", lambda: True)
        assert dialogs.pick_folder("随便") is None


class TestValidators:
    def test_validate_file(self, tmp_path):
        target = tmp_path / "a.txt"
        target.write_text("x", encoding="utf-8")
        assert dialogs.validate_file(target) is True
        assert dialogs.validate_file(tmp_path / "没有.txt") is False
        assert dialogs.validate_file(tmp_path) is False

    def test_validate_folder(self, tmp_path):
        assert dialogs.validate_folder(tmp_path) is True
        assert dialogs.validate_folder(tmp_path / "没有") is False


class TestAskPath:
    """``ask_path`` 是「弹窗 → 手输」的调度层，这里把两条路都走一遍。"""

    def test_uses_dialog_result_when_valid(self, tmp_path, monkeypatch):
        monkeypatch.setattr(dialogs, "dialog_available", lambda: True)
        picker = lambda **kwargs: str(tmp_path)  # noqa: E731
        result = dialogs.ask_path("folder", "提示：", picker, dialogs.validate_folder)
        assert result == str(tmp_path)

    def test_falls_back_to_console_when_dialog_cancelled(self, tmp_path, monkeypatch):
        monkeypatch.setattr(dialogs, "dialog_available", lambda: True)
        answers = iter([str(tmp_path)])
        result = dialogs.ask_path(
            "folder",
            "提示：",
            lambda **kwargs: None,
            dialogs.validate_folder,
            input_func=lambda prompt: next(answers),
        )
        assert result == str(tmp_path)

    def test_falls_back_when_dialog_returns_invalid_path(self, tmp_path, monkeypatch):
        monkeypatch.setattr(dialogs, "dialog_available", lambda: True)
        answers = iter([str(tmp_path)])
        result = dialogs.ask_path(
            "folder",
            "提示：",
            lambda **kwargs: str(tmp_path / "不存在"),
            dialogs.validate_folder,
            input_func=lambda prompt: next(answers),
        )
        assert result == str(tmp_path)

    def test_retries_until_valid(self, tmp_path, monkeypatch, capsys):
        monkeypatch.setattr(dialogs, "dialog_available", lambda: False)
        answers = iter([str(tmp_path / "不存在"), str(tmp_path)])
        result = dialogs.ask_path(
            "folder",
            "提示：",
            lambda **kwargs: None,
            dialogs.validate_folder,
            input_func=lambda prompt: next(answers),
        )
        assert result == str(tmp_path)
        assert "请重新输入" in capsys.readouterr().out

    def test_returns_none_on_empty_answer(self, monkeypatch, capsys):
        monkeypatch.setattr(dialogs, "dialog_available", lambda: False)
        result = dialogs.ask_path(
            "folder",
            "提示：",
            lambda **kwargs: None,
            dialogs.validate_folder,
            input_func=lambda prompt: "",
        )
        assert result is None
        assert "已取消" in capsys.readouterr().out

    def test_strips_quotes_and_whitespace(self, tmp_path, monkeypatch):
        """Windows 用户会把路径连同引号一起拖进来，得能处理。"""
        monkeypatch.setattr(dialogs, "dialog_available", lambda: False)
        result = dialogs.ask_path(
            "folder",
            "提示：",
            lambda **kwargs: None,
            dialogs.validate_folder,
            input_func=lambda prompt: f'  "{tmp_path}"  ',
        )
        assert result == str(tmp_path)

    def test_returns_absolute_path(self, tmp_path, monkeypatch):
        monkeypatch.setattr(dialogs, "dialog_available", lambda: False)
        result = dialogs.ask_path(
            "folder",
            "提示：",
            lambda **kwargs: None,
            dialogs.validate_folder,
            input_func=lambda prompt: ".",
        )
        assert result is not None
        assert result == os.path.abspath(".")

    def test_announces_when_no_gui(self, monkeypatch, capsys):
        monkeypatch.setattr(dialogs, "dialog_available", lambda: False)
        dialogs.ask_path(
            "file",
            "提示：",
            lambda **kwargs: None,
            dialogs.validate_file,
            input_func=lambda prompt: "",
        )
        assert "没有图形界面" in capsys.readouterr().out

    def test_passes_initial_dir_to_picker(self, monkeypatch, tmp_path):
        monkeypatch.setattr(dialogs, "dialog_available", lambda: True)
        seen = {}

        def picker(**kwargs):
            seen.update(kwargs)
            return None

        dialogs.ask_path(
            "folder",
            "提示：",
            picker,
            dialogs.validate_folder,
            initial_dir=tmp_path,
            input_func=lambda prompt: "",
        )
        assert seen["initial_dir"] == tmp_path


class TestFileTypePresets:
    def test_presets_are_tk_style_pairs(self):
        for preset in (dialogs.FILTER_STEP, dialogs.FILTER_PDF, dialogs.FILTER_EXCEL):
            assert preset
            for label, pattern in preset:
                assert label and isinstance(pattern, str)

    def test_step_filter_covers_both_extensions(self):
        pattern = dict(dialogs.FILTER_STEP)["STEP 模型"]
        assert "*.stp" in pattern and "*.step" in pattern


@pytest.mark.gui
@pytest.mark.skipif(not dialogs.dialog_available(), reason="需要图形界面")
def test_gui_manual_dialog_smoke():
    """手工用例：真的弹一次窗，确认不报错。默认不会被无人值守的流水线执行。"""
    pytest.skip("需要人工点窗口，默认跳过")
