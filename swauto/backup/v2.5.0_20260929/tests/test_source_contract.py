"""
测试组 A：源码一致性守护（guard）。

这组测试**不检查业务行为**，只检查一件事：
    src/*.bas 里那些被 tests/vba.py 重放的关键函数的**结构特征**是否还在。
一旦 VBA 侧被改动（重命名、常量换值、分支顺序颠倒、边界条件松动），
这里会立刻报错，提醒你同步更新 tests/vba.py —— 否则后面那些行为测试
就在测一个已经不存在的实现，变成假绿灯。

这组的价值：它是"测试自己会不会腐烂"的哨兵。
"""

from __future__ import annotations

import re

import pytest

from vba import body_of, extract_proc, load_source


@pytest.fixture(scope="module")
def util():
    return load_source("20_util.bas")


@pytest.fixture(scope="module")
def gather():
    return load_source("35_gather.bas")


@pytest.fixture(scope="module")
def dialog():
    return load_source("25_dialog.bas")


@pytest.fixture(scope="module")
def config():
    return load_source("10_config.bas")


# ---- 函数是否还都在（改名 = 立刻发现）----

@pytest.mark.parametrize(
    "proc",
    [
        "U_NormPath", "U_JoinPath", "U_FileName", "U_Parent", "U_BaseName",
        "U_SafeName", "U_Ext", "U_IsAbsPath", "U_FileKey",
        "U_Csv", "U_PsQuote",
        "U_HitKeyword", "U_HitBounded", "U_IsAsciiAlnum", "U_IsAsciiAlpha",
        "U_HasNonAscii", "U_HasSeparator", "U_KeywordList",
    ],
)
def test_util_procs_exist(util, proc):
    extract_proc(util, proc)


@pytest.mark.parametrize(
    "proc",
    ["G_KindSkipped", "G_SkipKind", "G_CompBadState", "G_StageSkipped",
     "G_MergeSkipped", "G_MergeOne", "G_StagedCount"],
)
def test_gather_procs_exist(gather, proc):
    extract_proc(gather, proc)


@pytest.mark.parametrize("proc", ["U_CleanResult", "U_AskSkipOptions"])
def test_dialog_procs_exist(dialog, proc):
    extract_proc(dialog, proc)


# ---- 关键边界条件是否还在（这些是历史 bug 的化身）----

def test_normpath_keeps_drive_root(util):
    """U_NormPath 的循环条件是 Len(s) > 3 —— 这保证 'D:\\' 不被削成 'D:'。"""
    body = body_of(extract_proc(util, "U_NormPath"))
    assert re.search(r"Len\(s\)\s*>\s*3", body), (
        "U_NormPath 的 'Len(s) > 3' 变了！如果改成 >= 3，'D:\\' 会被削成 'D:'，"
        "盘根路径直接损坏。改这条必须同步 tests/vba.py 的 norm_path()。"
    )


def test_joinpath_handles_empty_dir(util):
    """空目录时 U_JoinPath 必须**原样返回文件名**，不能拼出 '\\name'。"""
    body = body_of(extract_proc(util, "U_JoinPath"))
    assert re.search(r"If\s+Len\(s\)\s*=\s*0\s+Then", body)


def test_isabspath_requires_length_3(util):
    """U_IsAbsPath 先卡 Len < 3，避免 'D:' 这种半截路径被当成绝对路径。"""
    body = body_of(extract_proc(util, "U_IsAbsPath"))
    assert re.search(r"If\s+Len\(sPath\)\s*<\s*3\s+Then\s+Exit\s+Function", body)


def test_isabspath_checks_drive_and_unc(util):
    """两个分支都必须存在：盘符（第 2 字符是冒号）与 UNC（前两字符是双斜杠）。"""
    body = body_of(extract_proc(util, "U_IsAbsPath"))
    assert re.search(r'Mid\$\(sPath,\s*2,\s*1\)\s*=\s*":"', body), "盘符分支丢了"
    assert re.search(r'Left\$\(sPath,\s*2\)\s*=\s*"\\\\"', body), "UNC 分支丢了"


def test_hitbounded_is_strict_left_lax_right(util):
    """
    U_HitBounded 的核心不变量：**左边界严格（不能是字母数字）、右边界只拦字母**。
    这条只要被改成"两边都严格"，Nut8 / GBT5782 这类紧贴数字的写法就会全部漏判。
    """
    body = body_of(extract_proc(util, "U_HitBounded"))
    assert re.search(
        r"Not\s+U_IsAsciiAlnum\(cL\)\s+And\s+Not\s+U_IsAsciiAlpha\(cR\)", body
    ), (
        "U_HitBounded 的边界条件不再是（左:非字母数字 且 右:非字母）——"
        "这一条动了，标准件过滤会大面积误伤或漏判。"
    )


def test_hitbounded_advances_instr(util):
    """多候选必须继续往后找（InStr(p+1)），不能只判第一个就放弃。"""
    body = body_of(extract_proc(util, "U_HitBounded"))
    assert re.search(r"InStr\(p\s*\+\s*1,", body), (
        "U_HitBounded 不再推进搜索位置 —— 那么 'Pin, Pin9' 这类字符串只会看第一个 Pin。"
    )


def test_hitkeyword_three_branches(util):
    """U_HitKeyword 的三分支判定顺序：非 ASCII 或 含分隔符 → 子串；否则词边界。"""
    body = body_of(extract_proc(util, "U_HitKeyword"))
    assert re.search(r"U_HasNonAscii\(sNeedle\)\s+Or\s+U_HasSeparator\(sNeedle\)", body), (
        "U_HitKeyword 的「中文 / 含分隔符走子串匹配」分支变了。"
        "这条是「内六角螺钉M8x30」能命中的关键。"
    )


def test_hasseparator_uses_isalnum(util):
    """U_HasSeparator 用 U_IsAsciiAlnum 取反，不是硬编码字符表。"""
    body = body_of(extract_proc(util, "U_HasSeparator"))
    assert re.search(r"Not\s+U_IsAsciiAlnum\(", body)


def test_ascw_negative_wrap_present(util):
    """
    U_IsAsciiAlnum / U_IsAsciiAlpha / U_HasNonAscii 三处都要有 AscW 负数回绕。
    少了它，源码里出现 U+8000 以上的字符时会判错。
    """
    for proc in ("U_IsAsciiAlnum", "U_IsAsciiAlpha", "U_HasNonAscii"):
        body = body_of(extract_proc(util, proc))
        assert re.search(r"If\s+u\s*<\s*0\s+Then\s+u\s*=\s*u\s*\+\s*65536", body), (
            f"{proc} 丢了 AscW 的负数回绕（u < 0 → u += 65536）"
        )


def test_compbadstate_order_is_fixed(gather):
    """
    G_CompBadState 的判定顺序**必须**是 封套 → 压缩 → 隐藏。
    调换顺序会让同一个组件归到别的类别，勾选框里的计数和用户的预期就对不上了。
    """
    body = body_of(extract_proc(gather, "G_CompBadState"))
    i_env = body.find("IsEnvelope")
    i_sup = body.find("IsSuppressed")
    i_hid = body.find("IsHidden")
    assert -1 < i_env < i_sup < i_hid, (
        "G_CompBadState 的三类判定顺序变了（应严格为 封套 → 压缩 → 隐藏）"
    )


def test_ishidden_must_pass_false(gather):
    """
    【历史大坑】IsHidden 必须传 False。
    传 True 会把"轻化"组件一并当成隐藏 → 轻化零件被全部误杀。
    """
    body = body_of(extract_proc(gather, "G_CompBadState"))
    assert re.search(r"IsHidden\(False\)", body), (
        "IsHidden 不再传 False 了！传 True 会把轻化组件误判为隐藏并跳过 —— "
        "这会让轻化零件从交付物里凭空消失。"
    )
    assert not re.search(r"IsHidden\(True\)", body)


def test_skipkind_has_cycle_guard(gather):
    """G_SkipKind 的父级链回溯必须有 n > 64 的防环 bail，否则父级成环会挂死 SolidWorks。"""
    body = body_of(extract_proc(gather, "G_SkipKind"))
    assert re.search(r"If\s+n\s*>\s*64\s+Then\s+Exit\s+Do", body), (
        "G_SkipKind 的防环保护（n > 64）不见了 —— 父级链一旦成环，"
        "这个循环永远不结束，SolidWorks 会卡死。"
    )


def test_skipkind_clears_kind_on_success(gather):
    """整条链都正常时必须把 nKind 归零（等于显式声明"不跳过"），不能靠 VB 默认值。"""
    body = body_of(extract_proc(gather, "G_SkipKind"))
    assert re.search(r"Loop\s*\n\s*nKind\s*=\s*0", body)


def test_mergeone_three_branches(gather):
    """
    G_MergeOne 必须区分三种情况，顺序不能变：
      ① mask 命中 → 放回；② 没命中但路径已在待导出列表 → 什么都不写；③ 其余 → 记跳过。
    少了 ②，同一零件会出现"一边成功、一边跳过"的矛盾记录。
    """
    body = body_of(extract_proc(gather, "G_MergeOne"))
    assert re.search(r"If\s+\(nMask\s+And\s+nBit\)\s*<>\s*0", body), "分支① 丢了"
    assert re.search(r"ElseIf\s+Len\(sP\)\s*>\s*0\s+And\s+F_Contains\(", body), (
        "分支② 丢了 —— 这正是「同一零件不写矛盾记录」的除法。"
    )
    assert re.search(r"G_NoteSkip", body), "分支③ 丢了"


def test_mergeone_reports_unpathable(gather):
    """勾了但拿不到磁盘路径时必须如实写明，不能静默吞掉（用户会以为已导出）。"""
    body = body_of(extract_proc(gather, "G_MergeOne"))
    assert "取不到文件路径" in body


def test_cleanresult_strips_crlf_tab_bom(dialog):
    """U_CleanResult 的四个剔除动作一个都不能少（这是"取消变 C 却停不下来"的病根）。"""
    body = body_of(extract_proc(dialog, "U_CleanResult"))
    assert re.search(r'Replace\(s,\s*vbCr,\s*""\)', body), "vbCr 没剔"
    assert re.search(r'Replace\(s1,\s*vbLf,\s*""\)', body), "vbLf 没剔"
    assert re.search(r'Replace\(s1,\s*vbTab,\s*""\)', body), "vbTab 没剔"
    assert re.search(r"u\s*=\s*&HFEFF", body), "BOM 没剔"


def test_askskipoption_cancel_checked_before_length(dialog):
    """
    【历史大坑】U_AskSkipOptions 里 '"C"' 的判定**必须**排在 '"长度=3"' 之前。
    顺序反了，"取消导出"会退化成"没拿到结果"，用户按了取消却照旧导出。
    v2.2.0~v2.2.2 一直是这个 bug。
    """
    body = body_of(extract_proc(dialog, "U_AskSkipOptions"))
    i_cancel = body.find('If sResp = "C"')
    i_len = body.find("ElseIf Len(sResp) = 3")
    assert i_cancel != -1, "取消分支不见了"
    assert i_len != -1, "掩码分支不见了"
    assert i_cancel < i_len, (
        '顺序被调换了！\'If sResp = "C"\' 必须排在 \'ElseIf Len(sResp) = 3\' 之前。'
    )


def test_cfg_ver_is_single_source(config):
    """版本号只应有 CFG_VER 一处定义（其它地方都引用它）。"""
    hits = re.findall(r'CFG_VER\s+As\s+String\s*=\s*"([^"]+)"', config)
    assert len(hits) == 1, f"CFG_VER 定义了 {len(hits)} 次"
    assert re.fullmatch(r"v\d+\.\d+\.\d+", hits[0]), f"版本号格式不对: {hits[0]}"


def test_no_hardcoded_version_in_comments():
    """
    源码里**不许**再出现写死的版本号字面量。

    背景：00_header.bas 的文件头曾写死 "SWautoExport v2.2.0"，之后一路漂到 v2.4.2
    都没人发现；50_main.bas 的一行注释写死 "v2.0.2"，同样漂了好几个版本。
    现在两处都改成了引用 CFG_VER / 明确说明"不写版本号"，
    这条测试负责让它们别再长回来。

    注意：**例外是允许的** —— 注释里用版本号做"历史事故溯源"（例如
    "v2.2.0 到 v2.2.2 一直是这个毛病"）是必要的，那是在说明某段代码为什么长这样。
    这类引用会紧跟一个"历史叙述"的语境。所以这里只拦**声明式**的写法：
    形如 "SWautoExport v2.x.y" / "SWauto v2.x.y  某标题" 这种"当版本号在用"的写法。
    """
    offenders = []
    for name in ("00_header.bas", "10_config.bas", "20_util.bas", "25_dialog.bas",
                 "30_finder.bas", "35_gather.bas", "40_exporter.bas", "50_main.bas"):
        src = load_source(name)
        for i, ln in enumerate(src.split("\n"), 1):
            # 声明式用法：模块名/程序名 + 空格 + 版本号
            if re.search(r"SWauto(Export)?\s+v\d+\.\d+\.\d+", ln):
                # 允许："历史上 00_header.bas 的 v2.2.0 漂了..." 这类带引号的叙述
                if "漂" in ln or "历史上" in ln or "以前" in ln:
                    continue
                offenders.append(f"{name}:{i}: {ln.strip()}")
    assert not offenders, (
        "源码里又出现了写死的版本号（应该引用 CFG_VER）：\n  " + "\n  ".join(offenders)
    )


def test_header_has_no_version_line():
    """00_header.bas 的文件头不许再挂版本号（它漂 5 个版本那件事的补丁）。"""
    src = load_source("00_header.bas")
    head = "\n".join(src.split("\n")[:12])
    assert not re.search(r"SWautoExport\s+v\d+\.", head), (
        "00_header.bas 文件头又写死版本号了 —— 应该让它引用 CFG_VER，或干脆不写。"
    )


def test_kind_bits_are_1_2_4(gather):
    """三类跳过必须用 1/2/4 三个独立位 —— 用 1/2/3 会让"压缩"和"封套"撞车。"""
    body = body_of(extract_proc(gather, "G_CompBadState"))
    assert re.search(r"G_CompBadState\s*=\s*1", body)
    assert re.search(r"G_CompBadState\s*=\s*2", body)
    assert re.search(r"G_CompBadState\s*=\s*4", body)


# ---- COM 对象复用（U_Fso）----

def test_filesystemobject_is_cached(util):
    """
    U_FileExists / U_DirExists 不许每次调用都 CreateObject。
    装配体展开时"文件是否存在"要被问成千上万次，每次一个 COM 激活是真金白银的开销。
    正确的写法是走 U_Fso() 懒加载缓存。
    """
    for proc in ("U_FileExists", "U_DirExists"):
        body = body_of(extract_proc(util, proc))
        assert "U_Fso()" in body, f"{proc} 不再复用缓存的 FSO 了"
        assert "CreateObject" not in body, (
            f"{proc} 里又出现了直接 CreateObject —— 应该走 U_Fso() 复用。"
        )


def test_fso_lazy_singleton_pattern(util):
    """U_Fso 必须是"没有才建"的懒加载，而不是每次新建。"""
    body = body_of(extract_proc(util, "U_Fso"))
    assert re.search(r"If\s+m_fso\s+Is\s+Nothing\s+Then", body), (
        "U_Fso 丢了「已存在就不再建」的判断 —— 那就退化成每次 CreateObject 了"
    )
    assert re.search(r'CreateObject\("Scripting\.FileSystemObject"\)', body)


def test_m_fso_declared_in_header():
    """m_fso 必须声明在 00_header.bas（模块级），否则跨过程复用不了。"""
    header = load_source("00_header.bas")
    assert re.search(r"Private\s+m_fso\s+As\s+Object", header), (
        "00_header.bas 里没有 m_fso 的模块级声明"
    )


def test_cleanup_is_symmetric_with_init():
    """
    M_Cleanup 必须把 M_Init 建立/赋值的状态逐项清回去。
    早先只清了几个 Collection + m_SrcName，其余状态留着上一次运行的值 ——
    这种"部分清理"是那种将来加了新状态就必然忘记同步的写法。
    """
    main = load_source("50_main.bas")
    init = body_of(extract_proc(main, "M_Init"))
    cleanup = body_of(extract_proc(main, "M_Cleanup"))

    # 从 M_Init 里取出所有被赋值的模块级状态名
    assigned = set(re.findall(r"^\s*(?:Set\s+)?(m_\w+)\s*=", init, re.M))
    assert assigned, "M_Init 里没找到任何 m_* 赋值，抽取逻辑可能失效"

    missing = [name for name in sorted(assigned) if name not in cleanup]
    assert not missing, (
        "M_Init 赋值了但 M_Cleanup 没清的状态：" + ", ".join(missing) +
        "\n（新加状态时要两边同时改）"
    )


def test_cleanup_covers_startup_assigned_state():
    """
    补刀：M_Init 之外，各入口函数也会设一些状态（典型是 m_ModeTag），
    它们同样要出现在 M_Cleanup 里。上面那条测试只能看到 M_Init 赋值的，
    抓不到这类"由别处设置、但生命周期同属一次运行"的状态 —— 这条把它们钉住。
    """
    main = load_source("50_main.bas")
    cleanup = body_of(extract_proc(main, "M_Cleanup"))
    for name in ("m_ModeTag", "m_OutRoot", "m_StepDir", "m_PdfDir", "m_OriginalActive"):
        assert name in cleanup, (
            f"M_Cleanup 没清 {name} —— 它虽然在入口函数里赋值，"
            "但生命周期同样属于「一次运行」，不清会带到下一次。"
        )
