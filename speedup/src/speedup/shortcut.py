"""创建 Windows 快捷方式（``.lnk``）—— 纯 Python + ctypes 调 COM，**不写 .bat / .ps1 / .vbs**。

为什么要这个东西
----------------
图形界面光有 ``uv run speedup`` 还不够 —— 那还是得开命令行。要「双击就能开」，
需要一个指向 ``pythonw.exe``（无控制台窗口）的快捷方式。

试过但不可行的方案
------------------
* ``.pyw`` 文件：看着最省事，但它**依赖文件关联**。本机实测 ``HKEY_CLASSES_ROOT\\\\.pyw``
  没有默认值，双击只会弹出「你要如何打开这个文件」。快捷方式则是把 exe、参数、
  工作目录、图标全部写死，不依赖任何关联。
* ``.bat`` / ``.vbs`` 包一层：违反项目硬约束 R1（全 Python，不落 bat/psh/vbs）。

所以走 Windows 官方的方式：``IShellLinkW`` + ``IPersistFile``。这里用 ctypes
直接调 COM 虚表，因此**不需要 pywin32**，项目里不多一个依赖。

只调了六个方法
--------------
``SetPath`` / ``SetArguments`` / ``SetWorkingDirectory`` / ``SetDescription`` /
``SetIconLocation`` / ``IPersistFile::Save``。虚表下标是 MS-SHLLINK 与
``shobjidl.h`` 里定义的固定顺序，所以在模块级用常量写死并注明来源。
"""

from __future__ import annotations

import ctypes
import os
import sys
from ctypes import POINTER, Structure, byref, c_int, c_long, c_void_p, c_wchar_p
from dataclasses import dataclass
from pathlib import Path

from .config import PROJECT_ROOT
from .errors import SpeedupError
from .osutil import desktop_dir, is_windows

__all__ = [
    "SHORTCUT_NAME",
    "ShortcutSpec",
    "create_shortcut",
    "default_destination",
    "describe",
    "icon_path",
    "plan_shortcut",
    "shortcut_path",
]

#: 快捷方式的显示名（不含 .lnk 后缀）
SHORTCUT_NAME = "启动 speedup"

_CLSID_SHELL_LINK = "{00021401-0000-0000-C000-000000000046}"
_IID_SHELL_LINK_W = "{000214F9-0000-0000-C000-000000000046}"
_IID_PERSIST_FILE = "{0000010B-0000-0000-C000-000000000046}"

_CLSCTX_INPROC_SERVER = 1
_COINIT_APARTMENTTHREADED = 2

# IUnknown 的三个方法占虚表 0/1/2，之后才是接口自己的方法
_VT_RELEASE = 2

# IShellLinkW：GetPath(3) GetIDList(4) SetIDList(5) GetDescription(6) SetDescription(7)
#              GetWorkingDirectory(8) SetWorkingDirectory(9) GetArguments(10) SetArguments(11)
#              GetHotkey(12) SetHotkey(13) GetShowCmd(14) SetShowCmd(15)
#              GetIconLocation(16) SetIconLocation(17) SetRelativePath(18) Resolve(19) SetPath(20)
_VT_SET_DESCRIPTION = 7
_VT_SET_WORKING_DIRECTORY = 9
_VT_SET_ARGUMENTS = 11
_VT_SET_ICON_LOCATION = 17
_VT_SET_PATH = 20

# IPersistFile：GetClassID(3) IsDirty(4) Load(5) Save(6) SaveCompleted(7) GetCurFile(8)
_VT_QUERY_INTERFACE = 0
_VT_SAVE = 6


# ------------------------------------------------------------------ COM 底层


class _GUID(Structure):
    _fields_ = (
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    )


def _load_ole32():
    """取 ole32 并**显式声明参数类型** —— 64 位下指针对不上会直接崩，不能靠默认推断。"""
    try:
        ole32 = ctypes.windll.ole32  # type: ignore[attr-defined]
    except AttributeError as exc:  # pragma: no cover - 非 Windows
        raise SpeedupError("创建快捷方式只能在 Windows 上进行。") from exc

    ole32.CLSIDFromString.argtypes = [c_wchar_p, POINTER(_GUID)]
    ole32.CLSIDFromString.restype = c_long
    ole32.CoInitializeEx.argtypes = [c_void_p, ctypes.c_ulong]
    ole32.CoInitializeEx.restype = c_long
    ole32.CoUninitialize.argtypes = []
    ole32.CoUninitialize.restype = None
    ole32.CoCreateInstance.argtypes = [
        POINTER(_GUID),
        c_void_p,
        ctypes.c_ulong,
        POINTER(_GUID),
        POINTER(c_void_p),
    ]
    ole32.CoCreateInstance.restype = c_long
    return ole32


def _guid(text: str) -> _GUID:
    guid = _GUID()
    if _load_ole32().CLSIDFromString(c_wchar_p(text), byref(guid)) < 0:
        raise SpeedupError(f"GUID 字符串解析失败：{text}")
    return guid


def _method(pointer: c_void_p, index: int, *argtypes):
    """把 COM 对象虚表里第 ``index`` 个方法取出来，包成可调的 Python 函数。

    返回的函数第一个参数固定是 ``this`` 指针。
    """
    vtable = ctypes.cast(pointer, POINTER(POINTER(c_void_p))).contents
    address = vtable[index]
    if not address:  # pragma: no cover - 正常的 COM 对象不会有空槽
        raise SpeedupError(f"COM 虚表第 {index} 项是空的。")
    return ctypes.WINFUNCTYPE(c_long, c_void_p, *argtypes)(address)


def _release(pointer: c_void_p) -> None:
    if not pointer:
        return
    try:
        _method(pointer, _VT_RELEASE)(pointer)
    except Exception:  # pragma: no cover - 释放失败不该盖住真正的错误
        pass


def _check(result: int, what: str) -> None:
    if result < 0:
        raise SpeedupError(f"{what}失败（HRESULT 0x{result & 0xFFFFFFFF:08X}）")


def _write_lnk(spec: ShortcutSpec, destination: str) -> None:
    ole32 = _load_ole32()
    hr = ole32.CoInitializeEx(None, _COINIT_APARTMENTTHREADED)
    # S_FALSE（已经初始化过）和 RPC_E_CHANGED_MODE（别的线程模型）都不致命
    should_uninitialize = hr >= 0

    link = c_void_p()
    persist = c_void_p()
    try:
        clsid = _guid(_CLSID_SHELL_LINK)
        iid_link = _guid(_IID_SHELL_LINK_W)
        iid_persist = _guid(_IID_PERSIST_FILE)

        _check(
            ole32.CoCreateInstance(
                byref(clsid), None, _CLSCTX_INPROC_SERVER, byref(iid_link), byref(link)
            ),
            "创建 IShellLink 对象",
        )

        _check(
            _method(link, _VT_SET_PATH, c_wchar_p)(link, spec.target),
            "设置目标程序",
        )
        if spec.arguments:
            _check(
                _method(link, _VT_SET_ARGUMENTS, c_wchar_p)(link, spec.arguments),
                "设置启动参数",
            )
        if spec.working_dir:
            _check(
                _method(link, _VT_SET_WORKING_DIRECTORY, c_wchar_p)(link, spec.working_dir),
                "设置工作目录",
            )
        if spec.description:
            _check(
                _method(link, _VT_SET_DESCRIPTION, c_wchar_p)(link, spec.description),
                "设置备注",
            )
        if spec.icon:
            _check(
                _method(link, _VT_SET_ICON_LOCATION, c_wchar_p, c_int)(link, spec.icon, 0),
                "设置图标",
            )

        _check(
            _method(link, _VT_QUERY_INTERFACE, POINTER(_GUID), POINTER(c_void_p))(
                link, byref(iid_persist), byref(persist)
            ),
            "获取 IPersistFile 接口",
        )
        _check(
            _method(persist, _VT_SAVE, c_wchar_p, c_int)(persist, destination, 1),
            "写入 .lnk 文件",
        )
    finally:
        _release(persist)
        _release(link)
        if should_uninitialize:
            ole32.CoUninitialize()


# ------------------------------------------------------------------ 快捷方式内容


@dataclass(frozen=True)
class ShortcutSpec:
    """一个待创建的快捷方式。

    :param name: 显示名（不含 ``.lnk``）
    :param target: 要执行的 exe 绝对路径
    :param arguments: 传给 exe 的命令行参数
    :param working_dir: 起始目录
    :param icon: 图标文件；``None`` 表示用 exe 自带的图标
    :param description: 鼠标悬停时显示的备注
    """

    name: str
    target: str
    arguments: str
    working_dir: str
    icon: str | None
    description: str

    @property
    def filename(self) -> str:
        return f"{self.name}.lnk"


def icon_path() -> Path:
    """打包在包里的图标（``examples/make_icon.py`` 生成）。"""
    return Path(__file__).resolve().parent / "assets" / "speedup.ico"


def _pythonw() -> str:
    """挑一个 ``pythonw.exe`` —— 它启动的进程没有控制台黑框。

    依次尝试：当前解释器旁边 → 项目 venv 里。都没有就退回当前解释器
    （会有黑框，但功能不受影响；宁可难看也别不能用）。
    """
    for candidate in (
        Path(sys.executable).with_name("pythonw.exe"),
        PROJECT_ROOT / ".venv" / "Scripts" / "pythonw.exe",
    ):
        if candidate.is_file():
            return str(candidate)
    return sys.executable


def plan_shortcut(name: str = SHORTCUT_NAME) -> ShortcutSpec:
    """按当前项目算出快捷方式该怎么配。纯计算，不落盘。"""
    icon = icon_path()
    return ShortcutSpec(
        name=name,
        target=_pythonw(),
        arguments="-m speedup",
        working_dir=str(PROJECT_ROOT),
        icon=str(icon) if icon.is_file() else None,
        description="打开 speedup 车间提速工具集的图形界面（双击即用，不用记命令）",
    )


def default_destination(*, desktop: bool = False) -> Path:
    """默认保存位置：项目根目录，或桌面。"""
    if not desktop:
        return PROJECT_ROOT
    found = desktop_dir()
    if found is None:
        raise SpeedupError("找不到桌面目录，请用 --dest 指定保存位置。")
    return Path(found)


def shortcut_path(spec: ShortcutSpec, directory: str | os.PathLike[str]) -> Path:
    """算出快捷方式会落在哪个文件上。纯函数，不落盘。"""
    folder = Path(directory)
    if not folder.is_dir():
        raise SpeedupError(f"保存目录不存在：{folder}")
    return folder / spec.filename


def create_shortcut(spec: ShortcutSpec, path: str | os.PathLike[str]) -> Path:
    """把快捷方式写到 ``path``（完整文件路径，用 :func:`shortcut_path` 算）。

    返回实际写入的路径。
    """
    if not is_windows():  # pragma: no cover - 本机是 Windows
        raise SpeedupError("快捷方式只能在 Windows 上创建。")
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    _write_lnk(spec, str(target))
    if not target.is_file():  # pragma: no cover - COM 报成功却没落盘，属于异常情况
        raise SpeedupError(f"写入后没找到文件，可能被杀毒软件拦了：{target}")
    return target


def describe(spec: ShortcutSpec) -> str:
    """把人需要核对的信息排成几行，供 CLI 打印。"""
    lines = [
        f"名称      {spec.filename}",
        f"目标程序  {spec.target}",
        f"启动参数  {spec.arguments}",
        f"起始位置  {spec.working_dir}",
        f"图标      {spec.icon or '（用目标程序自带图标）'}",
        f"备注      {spec.description}",
    ]
    return "\n".join(lines)
