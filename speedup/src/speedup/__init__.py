"""speedup —— 车间提速工具集。

纯 Python 实现，**不依赖 bat / PowerShell**。所有交互（文件选择、文件夹选择）
走 tkinter，tkinter 不可用时回退控制台输入。

对外只暴露版本号；具体工具在 :mod:`speedup.tools`，命令行入口在 :mod:`speedup.cli`。
"""

from __future__ import annotations

__version__ = "0.1.0"
__all__ = ["__version__"]
