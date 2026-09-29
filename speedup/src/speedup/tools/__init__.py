"""工具注册表。

每个工具有自己的模块，统一暴露四个东西：

* ``NAME`` / ``TITLE`` / ``SUMMARY`` —— 供 CLI 子命令与交互菜单展示
* ``add_parser(subparsers)`` —— 注册 argparse 子命令
* ``run(args)`` —— 执行，返回进程退出码（0 = 成功）

:mod:`speedup.cli` 只依赖这张表，新增工具时**只需要往 :data:`TOOL_MODULES`
里加一行**，菜单和子命令会自动出现。
"""

from __future__ import annotations

from types import ModuleType

from . import gbbuild, gbcopy

__all__ = ["TOOL_MODULES", "TOOLS"]


def _sorted_tools() -> tuple[ModuleType, ...]:
    """按 NAME 排序，保证菜单顺序稳定（不依赖 import 顺序）。"""
    return tuple(sorted(TOOL_MODULES, key=lambda module: module.NAME))


TOOL_MODULES: tuple[ModuleType, ...] = (gbbuild, gbcopy)

#: ``{子命令名: 模块}``
TOOLS: dict[str, ModuleType] = {module.NAME: module for module in _sorted_tools()}
