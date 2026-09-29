"""
VBA 语义模拟层 —— 给 SWauto 的纯逻辑函数在 Python 里做等价重放。

设计原则（很重要，别改成"抄一份 Python 实现"）：
    SWauto 的真正实现是 src/*.bas 里的 VBA。如果测试里再手写一份 Python 抄本，
    那 VBA 改了、抄本没改，测试照样绿 —— 这种测试是负资产。
    所以这里的做法是：**从 .bas 源码里把待测逻辑原样抽出来**，
    再用本模块提供的 VBA 内建函数（AscW / Mid$ / InStrRev / UCase$ …）
    在 Python 里跑一遍。源码一改，抽出来的表达式就变了，测试立刻能发现。

本模块只负责"让 VBA 内建函数在 Python 里长得一样"，不含任何业务逻辑。
"""

from __future__ import annotations

import re
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent.parent / "src"


# --------------------------------------------------------------------------
# VBA 内建函数的等价实现
# 命名刻意带尾巴下划线，避免和 Python 内建名混淆。
# --------------------------------------------------------------------------

def load_source(module: str) -> str:
    """按 UTF-8 读 src/<module>（SWauto 源码是 UTF-8 + LF，无 BOM）。"""
    return (SRC_DIR / module).read_text(encoding="utf-8")


def extract_proc(source: str, name: str) -> str:
    """
    从源码里抽取一个 Sub/Function 的正文（含签名与 End Xxx），不含前导注释块。

    抽的是"逻辑行"（先做行续接合并），所以返回的文本里没有 VB 的 `_` 续行符，
    可以当成单行长表达式处理。
    """
    lines = source.split("\n")
    # 先合并续行，并保留原始行号信息（这里只关心正文，不做行号映射）
    logical: list[str] = []
    buf = ""
    for ln in lines:
        s = ln.rstrip()
        if s.endswith("_"):
            buf += s[:-1] + " "
        else:
            logical.append(buf + s)
            buf = ""
    if buf:
        logical.append(buf)

    start = None
    for i, ln in enumerate(logical):
        s = ln.strip()
        if re.match(rf"^(Private |Public |Friend )?(Static )?(Sub|Function)\s+{re.escape(name)}\b", s):
            start = i
            break
    if start is None:
        raise LookupError(f"{name} 不在源码里（函数被改名或删了？）")

    is_sub = re.match(
        r"^(Private |Public |Friend )?(Static )?Sub\s+", logical[start].strip()
    )
    end_re = r"^End Sub\b" if is_sub else r"^End Function\b"

    for j in range(start + 1, len(logical)):
        if re.match(end_re, logical[j].strip()):
            return "\n".join(logical[start : j + 1])
    raise LookupError(f"{name} 没找到对应的 End Sub/End Function")


def body_of(proc: str) -> str:
    """取过程正文（去掉第一行签名和最后一行 End Xxx）。"""
    ls = proc.split("\n")
    return "\n".join(ls[1:-1])


# ---- VB 字符串函数 ----

def mid(s: str, start: int, length: int | None = None) -> str:
    """VB 的 Mid$/Mid：1-based，越界返回空串，绝不抛异常。"""
    if start < 1 or start > len(s):
        return ""
    if length is None:
        return s[start - 1 :]
    if length < 0:
        raise ValueError("Mid$ 的长度不能为负")
    return s[start - 1 : start - 1 + length]


def left(s: str, n: int) -> str:
    if n <= 0:
        return ""
    return s[:n]


def right(s: str, n: int) -> str:
    if n <= 0:
        return ""
    return s[-n:] if n <= len(s) else s


def instrrev(s: str, find: str, start: int | None = None) -> int:
    """VB 的 InStrRev：返回**从 1 计起**的位置，找不到返回 0。"""
    if start is None:
        start = len(s)
    # VB 允许 start = -1 表示结尾
    if start == -1:
        start = len(s)
    if start < 1 or start > len(s):
        return 0
    idx = s.rfind(find, 0, start + len(find) - 1 if find else start)
    if find == "":
        return 0
    idx = s.rfind(find, 0, start)
    if idx < 0:
        return 0
    return idx + 1


def instr(s: str, find: str) -> int:
    """VB 的 InStr(1, hay, needle, vbTextCompare)：1-based，找不到返回 0。"""
    return s.lower().find(find.lower()) + 1


def ascw(ch: str) -> int:
    """
    VBA 的 AscW：返回**有符号 16 位**值，U+8000 以上为负。
    源码里到处都在 `If u < 0 Then u = u + 65536` 把它掰回来 ——
    那个回绕行为必须在这里如实模拟，否则边界分支测不到。
    """
    if not ch:
        raise ValueError("AscW 需要恰好一个字符")
    u = ord(ch[0])
    if u > 0xFFFF:
        raise ValueError("VBA 的 AscW 只处理 BMP（U+FFFF 以内）")
    return u - 0x10000 if u >= 0x8000 else u


def ucase(s: str) -> str:
    return s.upper()


def vba_trim(s: str) -> str:
    """VBA 的 Trim$ 只去**空格**，不去 Tab / CR / LF —— 这点和 Python 的 strip() 不同。"""
    return s.strip(" ")


# --------------------------------------------------------------------------
# 模拟为"函数的等价实现"：这些是从源码抽出来后、由测试驱动的重放
# --------------------------------------------------------------------------

def is_ascii_alnum(c: str) -> bool:
    """对应 U_IsAsciiAlnum：空串 = False（源码 `If Len(c) = 0 Then Exit Function`）。"""
    if not c:
        return False
    u = ascw(c)
    if u < 0:
        u += 65536
    return (48 <= u <= 57) or (65 <= u <= 90) or (97 <= u <= 122)


def is_ascii_alpha(c: str) -> bool:
    """对应 U_IsAsciiAlpha。"""
    if not c:
        return False
    u = ascw(c)
    if u < 0:
        u += 65536
    return (65 <= u <= 90) or (97 <= u <= 122)


def has_non_ascii(s: str) -> bool:
    """对应 U_HasNonAscii：> 126 即算非 ASCII。"""
    for c in s:
        u = ascw(c)
        if u < 0:
            u += 65536
        if u > 126:
            return True
    return False


def has_separator(s: str) -> bool:
    """对应 U_HasSeparator：含任何**非字母数字**字符即为真。"""
    for c in s:
        if not is_ascii_alnum(c):
            return True
    return False


def hit_bounded(hay: str, needle: str) -> bool:
    """
    对应 U_HitBounded：纯字母数字关键字的"前边界严格、后边界放行数字"匹配。
    这是全项目最容易踩坑的一段，逐句照搬源码（含 `InStr(p + 1, ...)` 的推进方式）。
    """
    if not hay or not needle:
        return False
    n = len(needle)
    low_hay = hay.lower()
    low_needle = needle.lower()
    p = low_hay.find(low_needle)
    while p >= 0:
        c_l = hay[p - 1] if p > 0 else ""
        c_r = hay[p + n] if p + n < len(hay) else ""
        if not is_ascii_alnum(c_l) and not is_ascii_alpha(c_r):
            return True
        p = low_hay.find(low_needle, p + 1)
    return False


def hit_keyword(hay: str, needle: str) -> bool:
    """对应 U_HitKeyword 的三分支规则。"""
    if not hay or not needle:
        return False
    if has_non_ascii(needle) or has_separator(needle):
        return needle.lower() in hay.lower()
    return hit_bounded(hay, needle)


def keyword_list(csv: str) -> list[str]:
    """对应 U_KeywordList：按逗号切、去空、丢空项。"""
    return [x.strip(" ") for x in csv.split(",") if x.strip(" ")]


def safe_name(s: str) -> str:
    """对应 U_SafeName：Windows 非法文件名字符 → 下划线。"""
    out = s
    for ch in '\\/:*?"<>|':
        out = out.replace(ch, "_")
    return vba_trim(out)


# --------------------------------------------------------------------------
# 路径函数（对应 20_util.bas 第 2.1 节）
# 注意 U_NormPath 的去尾斜杠条件：`Len(s) > 3` —— 所以 "D:\" 会保留尾斜杠。
# --------------------------------------------------------------------------

def norm_path(s: str) -> str:
    s = vba_trim(s)
    while len(s) > 3 and s[-1] == "\\":
        s = s[:-1]
    return s


def join_path(d: str, name: str) -> str:
    s = norm_path(d)
    if len(s) == 0:
        return name
    if s[-1] != "\\":
        s += "\\"
    return s + name


def file_name(path: str) -> str:
    i = instrrev(path, "\\")
    return path if i == 0 else mid(path, i + 1)


def parent(path: str) -> str:
    i = instrrev(path, "\\")
    if i <= 0:
        return ""
    if i == 3:
        return left(path, 3)
    return left(path, i - 1)


def base_name(path: str) -> str:
    s = file_name(path)
    i = instrrev(s, ".")
    if i > 1:
        s = left(s, i - 1)
    return s


def ext(path: str) -> str:
    s = file_name(path)
    i = instrrev(s, ".")
    return ucase(mid(s, i)) if i > 0 else ""


def is_abs_path(path: str) -> bool:
    """对应 U_IsAbsPath。空串/短串走 VB 默认 False。"""
    if len(path) < 3:
        return False
    if mid(path, 2, 1) == ":":
        return True
    return left(path, 2) == "\\\\"


def file_key(path: str) -> str:
    return ucase(norm_path(vba_trim(path)))


def is_model(path: str) -> bool:
    e = ext(path)
    return e in (".SLDPRT", ".SLDASM")


def is_drawing(path: str) -> bool:
    return ext(path) == ".SLDDRW"


# --------------------------------------------------------------------------
# 对话框结果清洗（对应 25_dialog.bas U_CleanResult）
# --------------------------------------------------------------------------

def clean_result(s: str) -> str:
    """CR / LF / Tab / 开头 BOM 全部剔掉，最后 Trim$（只去空格）。"""
    if not s:
        return ""
    s1 = s.replace("\r", "").replace("\n", "").replace("\t", "")
    if s1:
        u = ascw(left(s1, 1))
        if u < 0:
            u += 65536
        if u == 0xFEFF:
            s1 = mid(s1, 2)
    return vba_trim(s1)


def ps_quote(s: str) -> str:
    """对应 U_PsQuote：包单引号，内部单引号翻倍。"""
    return "'" + s.replace("'", "''") + "'"


def csv_cell(s: str) -> str:
    """对应 U_Csv：用双引号包，内部双引号翻倍。"""
    return '"' + s.replace('"', '""') + '"'


# --------------------------------------------------------------------------
# 跳过掩码（对应 35_gather.bas 的暂存 / 合并）
# --------------------------------------------------------------------------

def kind_bit(kind: int) -> int:
    """G_KindSkipped 用到的位：1 封套 / 2 已压缩 / 4 已隐藏。"""
    return kind


def kind_skipped(kind: int, cfg_env: bool, cfg_sup: bool, cfg_hid: bool) -> bool:
    """对应 G_KindSkipped —— CFG_SKIP_* 三开关按类别取值。"""
    if kind == 1:
        return cfg_env
    if kind == 2:
        return cfg_sup
    if kind == 4:
        return cfg_hid
    return False


def merge_one(rows, bit: int, mask: int, kind_name: str, models, results):
    """
    对应 G_MergeOne。rows 每项是 (label, path, why)。
    逐句照搬它的三分支：
      ① 勾了这类（mask 含 bit）→ 有路径就放回待导出；没路径写明"取不到路径"
      ② 没勾、但该路径已在待导出列表里 → 什么都不写（避免自相矛盾）
      ③ 其余 → 记一条跳过
    """
    for label, path, why in rows:
        if mask & bit:
            if len(path) > 0:
                if path.lower() not in [m.lower() for m in models]:
                    models.append(path)
            else:
                results.append(("模型", label, "", "跳过", f"{kind_name}，但取不到文件路径，无法导出"))
        elif len(path) > 0 and path.lower() in [m.lower() for m in models]:
            continue
        else:
            results.append(("SKIP", label, why))
    return models, results


# --------------------------------------------------------------------------
# 装配体展开的父级链遍历（对应 G_SkipKind 的防环逻辑）
# --------------------------------------------------------------------------

def skip_kind_chain(states, limit=64):
    """
    对应 G_SkipKind：沿父级链往上找第一个"状态不正常"的组件。
    states 是自身到最顶层的状态列表（0 正常 / 1 封套 / 2 压缩 / 4 隐藏）。
    返回 (kind, why, hops)；kind = 0 表示整条链都正常。
    """
    names = {1: "封套", 2: "已压缩", 4: "已隐藏"}
    n = 0
    for st in states:
        n += 1
        if n > limit:
            return 0, "", n - 1          # 防御性 bail，与源码 `Exit Do` 一致
        if st != 0:
            why = names.get(st, "")
            if n > 1:
                why += "（上级装配体）"
            return st, why, n
    return 0, "", min(n, limit)
