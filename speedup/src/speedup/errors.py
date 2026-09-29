"""项目自定义异常。

约定：所有「用户可理解的失败」都抛 :class:`SpeedupError` 的子类，
由 :mod:`speedup.cli` 统一捕获并打印成中文提示，不吐 traceback。
未预期的异常才让它冒泡（用于开发期定位 bug）。
"""

from __future__ import annotations

__all__ = [
    "SpeedupError",
    "UserCancelled",
    "UnsupportedFormatError",
    "InputError",
    "MissingDependencyError",
    "ConfigError",
]


class SpeedupError(Exception):
    """本工具集所有可预期错误的基类。"""


class UserCancelled(SpeedupError):
    """用户主动取消（关掉了选择窗口 / 输入了空行）。"""


class UnsupportedFormatError(SpeedupError):
    """输入文件格式不受支持，例如旧的 .xls。"""


class InputError(SpeedupError):
    """输入内容不符合预期，例如找不到「名称」列、文件夹里没有模型文件。"""


class MissingDependencyError(SpeedupError):
    """缺少可选依赖，例如没有安装 pdfplumber。"""


class ConfigError(SpeedupError):
    """配置文件缺失、格式错误或字段非法。"""
