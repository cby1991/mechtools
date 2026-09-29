"""pytest 公共夹具。

约定：测试里**不写业务魔数**，需要常量时从 :mod:`speedup.config` 取，
这样常量改了测试会跟着走，不会出现「测试还在断言旧的列号」。
"""

from __future__ import annotations

import os
from collections.abc import Sequence
from pathlib import Path

import pytest

from speedup.config import MODEL_EXTS
from speedup.matching import ScannedFile


@pytest.fixture
def workdir(tmp_path: Path) -> Path:
    """一个干净的临时工作目录。"""
    target = tmp_path / "work"
    target.mkdir()
    return target


@pytest.fixture
def make_files():
    """在目录下批量造空文件，返回路径列表。

    用法::

        >>> paths = make_files(tmp_path, ["a.step", "b.pdf"])
    """

    def _make(directory: str | os.PathLike[str], names: Sequence[str]) -> list[Path]:
        base = Path(directory)
        base.mkdir(parents=True, exist_ok=True)
        created = []
        for name in names:
            path = base / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(b"")
            created.append(path)
        return created

    return _make


@pytest.fixture
def scanned(make_files, tmp_path):
    """把一批文件名转换成 :class:`ScannedFile` 列表（不落盘也能用）。"""

    def _scanned(names: Sequence[str]) -> list[ScannedFile]:
        made = make_files(tmp_path / "candidates", names)
        return [
            ScannedFile(
                path=str(path),
                stem=path.stem,
                ext=path.suffix.lower(),
            )
            for path in made
        ]

    return _scanned


@pytest.fixture
def model_ext() -> str:
    return MODEL_EXTS[0]


@pytest.fixture(autouse=True)
def _no_gui_side_effects(monkeypatch):
    """避免测试期间误开资源管理器 / 默认程序（``os.startfile``）。

    只打工具的**使用点**，不动 :mod:`speedup.osutil` 本身 —— 否则
    ``test_osutil.py`` 就没法测真实实现（各工具是 ``from ..osutil import ...``
    导入的，持有自己的引用，所以必须在使用点打补丁）。
    """
    for target in (
        "speedup.tools.gbbuild.open_with_default_app",
        "speedup.tools.gbcopy.open_with_default_app",
    ):
        monkeypatch.setattr(target, lambda *args, **kwargs: True)
    yield
