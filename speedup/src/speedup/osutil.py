"""跨平台的小工具：打开文件、定位文件夹、判断 Windows。

原脚本直接调 ``os.startfile``，在非 Windows 上会 AttributeError。
这里统一包一层，失败只返回 ``False``，不让「顺手打开一下」这种附加功能
把主流程搞挂。
"""

from __future__ import annotations

import os
import subprocess
import sys

__all__ = ["is_windows", "open_with_default_app", "reveal_in_file_manager", "desktop_dir"]


def is_windows() -> bool:
    return sys.platform.startswith("win")


def desktop_dir() -> str | None:
    """用户桌面目录；找不到返回 ``None``。

    不能直接写 ``~/Desktop``：装了 OneDrive 的机器上桌面会被重定向到
    ``%USERPROFILE%\\OneDrive\\桌面``。所以优先读注册表里的真实位置
    （值可能是 ``%USERPROFILE%\\Desktop`` 这种带环境变量的形式，要展开）。
    """
    if is_windows():
        try:
            import winreg

            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
            ) as key:
                raw, _kind = winreg.QueryValueEx(key, "Desktop")
            expanded = os.path.expandvars(raw)
            if os.path.isdir(expanded):
                return expanded
        except (OSError, ImportError, ValueError):
            pass
    fallback = os.path.join(os.path.expanduser("~"), "Desktop")
    return fallback if os.path.isdir(fallback) else None


def open_with_default_app(path: str | os.PathLike[str]) -> bool:
    """用系统默认程序打开文件/文件夹。成功返回 ``True``。"""
    target = os.fspath(path)
    if is_windows():
        try:
            os.startfile(target)  # type: ignore[attr-defined]  # noqa: S606
        except OSError:
            return False
        return True
    opener = "open" if sys.platform == "darwin" else "xdg-open"
    try:
        subprocess.Popen([opener, target], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except OSError:
        return False
    return True


def reveal_in_file_manager(path: str | os.PathLike[str]) -> bool:
    """在文件管理器里定位到该路径（Windows 用 ``explorer /select,``）。"""
    target = os.fspath(path)
    if is_windows():
        try:
            os.startfile(target)  # type: ignore[attr-defined]  # noqa: S606
        except OSError:
            return False
        return True
    return open_with_default_app(os.path.dirname(os.path.abspath(target)))
