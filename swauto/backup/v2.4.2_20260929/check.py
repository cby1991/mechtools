import re

# dist 里的 .bas 是 GBK 编码（VBE 按系统 ANSI 代码页解码），这里也要按 GBK 读
ENC = "gbk"
src = open("dist/SWautoExport.bas", encoding=ENC).read()
lines = src.split("\n")   # 文本模式已把 CRLF 归一为 LF


def strip_comment(ln):
    """去掉字符串外的注释（' 之后的都算注释）"""
    out = []
    inq = False
    for c in ln:
        if c == '"':
            inq = not inq
            out.append(c)
        elif c == "'" and not inq:
            break
        else:
            out.append(c)
    return "".join(out)


code = [strip_comment(l).rstrip() for l in lines]

# 把 VB 的行续接（行尾 "_"）合并成逻辑行，否则多行 If / 多行 MsgBox 会被漏统计
logical = []
buf = ""
for ln in code:
    s = ln.rstrip()
    if s.endswith("_"):
        buf += s[:-1] + " "
    else:
        logical.append(buf + s)
        buf = ""
if buf:
    logical.append(buf)

counts = {"Sub": 0, "End Sub": 0, "Function": 0, "End Function": 0,
          "If_Then_block": 0, "End If": 0, "For": 0, "Next": 0,
          "Do": 0, "Loop": 0, "Select Case": 0, "End Select": 0}
quote_bad = []       # 引号不成对的行（吞引号 = 语法错误的典型症状）
for idx, ln in enumerate(logical, 1):
    s = ln.strip()
    if re.match(r"^(Private |Public |Friend )?(Static )?Sub \w+", s):
        counts["Sub"] += 1
    elif s.startswith("End Sub"):
        counts["End Sub"] += 1
    elif re.match(r"^(Private |Public |Friend )?(Static )?Function \w+", s):
        counts["Function"] += 1
    elif s.startswith("End Function"):
        counts["End Function"] += 1
    elif re.match(r"^If\b.*\bThen$", s):
        counts["If_Then_block"] += 1
    elif s.startswith("End If"):
        counts["End If"] += 1
    elif re.match(r"^(For|For Each)\b", s):
        counts["For"] += 1
    elif s.startswith("Next"):
        counts["Next"] += 1
    elif re.match(r"^Do\b", s):
        counts["Do"] += 1
    elif s.startswith("Loop"):
        counts["Loop"] += 1
    elif re.match(r"^Select Case\b", s):
        counts["Select Case"] += 1
    elif s.startswith("End Select"):
        counts["End Select"] += 1
    if ln.count('"') % 2 != 0:
        quote_bad.append(idx)

report = []
report.append("=== 读取编码: %s ===" % ENC.upper())
report.append("")
report.append("=== 过程块配平 ===")
# 「多行 If」的续行行尾也会是 Then，统计时会少算 1 个，属正常偏差
for a, b in [("Sub", "End Sub"), ("Function", "End Function"),
             ("If_Then_block", "End If"), ("For", "Next"),
             ("Do", "Loop"), ("Select Case", "End Select")]:
    mark = "OK  " if counts[a] == counts[b] else "*** 不平"
    report.append("%s %-14s %3d  /  %-14s %3d" % (mark, a, counts[a], b, counts[b]))

report.append("")
report.append("=== 引号成对性（吞引号会直接导致语法错误）===")
if quote_bad:
    for n in quote_bad:
        report.append("  *** 第 %d 行引号数为奇数: %s" % (n, lines[n - 1].strip()[:80]))
else:
    report.append("  全部行引号成对 OK")

report.append("")
# 用剥掉注释后的代码扫描，否则注释里提到的枚举名会被误算成"真依赖"
codelines = "\n".join(code)
report.append("=== 类型库依赖 ===")
for x in sorted(set(re.findall(r"\b(sw[A-Za-z]+_e)\b", codelines))):
    report.append("  " + x)
report.append("--- 枚举成员 ---")
for x in sorted(set(re.findall(r"\bsw[A-Za-z]+_e\.(sw[A-Za-z0-9_]+)", codelines))):
    report.append("  " + x)

report.append("")
report.append("=== 配置常量一致性（用了但没定义 = VBE 报「变量未定义」）===")
cfg_def = set(re.findall(r"Const\s+(CFG_\w+)", codelines))
cfg_use = set(re.findall(r"\b(CFG_\w+)\b", codelines))
cfg_missing = sorted(cfg_use - cfg_def)
cfg_unused = sorted(cfg_def - cfg_use)
if cfg_missing:
    for x in cfg_missing:
        report.append("  *** 使用了但未定义: " + x)
else:
    report.append("  全部 %d 个 CFG_* 均已定义 OK" % len(cfg_def))
if cfg_unused:
    report.append("  （定义了但没被引用，不影响运行，仅提示: %s）" % ", ".join(cfg_unused))

report.append("")
report.append("=== 标准件关键字规则抽检（读上方配置区的实际值）===")


def vb_string_const(lines_, name):
    """从 VB 源码里取出一个（可能用 _ 跨行续接的）字符串常量的值"""
    for i, ln in enumerate(lines_):
        if re.search(r"Const\s+%s\s+As\s+String\s*=" % name, ln):
            buf, j = "", i
            while j < len(lines_):
                buf += lines_[j]
                if lines_[j].rstrip().endswith("_"):
                    j += 1
                    continue
                break
            body = buf.split("=", 1)[1]
            return "".join(re.findall(r'"([^"]*)"', body))
    return ""


def _is_alnum(c):
    """与 VBA 里的 U_IsAsciiAlnum 保持一致：只有 ASCII 字母/数字算，中文不算"""
    if len(c) != 1:
        return False
    o = ord(c)
    return (48 <= o <= 57) or (65 <= o <= 90) or (97 <= o <= 122)


def _is_alpha(c):
    if len(c) != 1:
        return False
    o = ord(c)
    return (65 <= o <= 90) or (97 <= o <= 122)


def _has_non_ascii(s):
    return any(ord(c) > 126 for c in s)


def _has_separator(s):
    return any(not _is_alnum(c) for c in s)


def _hit_bounded(hay, needle):
    for m in re.finditer(re.escape(needle), hay, re.IGNORECASE):
        left = hay[m.start() - 1] if m.start() > 0 else ""
        right = hay[m.end()] if m.end() < len(hay) else ""
        # 前边界不能是字母数字；后边界不能是字母（数字放行，兼容 GBT5782 / ISO4762 / Nut8）
        if not _is_alnum(left) and not _is_alpha(right):
            return True
    return False


def hit_keyword(hay, needle):
    """Python 版镜像 VBA 的 U_HitKeyword（务必与 src/20_util.bas 保持一致）"""
    if not hay or not needle:
        return False
    if _has_non_ascii(needle) or _has_separator(needle):
        return needle.lower() in hay.lower()     # 中文 / 自带分隔符 → 子串匹配
    return _hit_bounded(hay, needle)             # 纯字母数字 → 词边界匹配


_kw = vb_string_const(code, "CFG_STD_KEYWORDS")
_codes = vb_string_const(code, "CFG_STD_STDCODES")
_paths = vb_string_const(code, "CFG_STD_PATH_HINTS")
report.append("  文件名关键字(%d): %s" % (len([x for x in _kw.split(",") if x.strip()]), _kw))
report.append("  标准号(%d)     : %s" % (len([x for x in _codes.split(",") if x.strip()]), _codes))
report.append("  路径关键字(%d) : %s" % (len([x for x in _paths.split(",") if x.strip()]), _paths))
report.append("")

_ws = [x.strip() for x in (_kw.split(",") + _codes.split(",")) if x.strip()]


def _judge(name):
    for k in _ws:
        if hit_keyword(name, k):
            return k
    return None


# 应该命中：注意「螺钉M8」「GBT5782」这类"关键字紧贴尺寸/代号"的写法，必须能认出来
SHOULD_HIT = ["螺母 M8", "螺母M8", "内六角螺钉M8x30", "GBT5782 六角头螺栓", "ISO4762",
              "GB-T 5782-2000 螺栓", "Hex Bolt, Fully Threaded", "O-Ring_10", "Spring Pin",
              "垫圈 GB97.1", "弹性挡圈 12", "DIN 912 内六角", "Split Washer 8", "销钉8x20",
              "Nut8", "开槽圆柱销 6x30"]

# 不该命中：全部是"关键字藏在别的词里面"的陷阱
SHOULD_MISS = ["Spindle", "Walnut", "Nutshell", "ISOLATION", "Screwdriver", "Bolted",
               "外壳", "上盖板", "主轴", "定位块", "Flange Plate", "Base Bracket",
               "轴承座", "导向板"]

_bad = []
for n in SHOULD_HIT:
    if not _judge(n):
        _bad.append("  *** 应该命中却漏了: %s" % n)
for n in SHOULD_MISS:
    hit = _judge(n)
    if hit:
        _bad.append("  *** 不该命中却中了(%s): %s" % (hit, n))
if _bad:
    report.extend(_bad)
else:
    report.append("  应该命中 %d/%d OK   不该命中 %d/%d OK"
                  % (len(SHOULD_HIT), len(SHOULD_HIT), len(SHOULD_MISS), len(SHOULD_MISS)))
    report.append("  规则：中文/带分隔符的关键字走子串匹配（螺钉M8 能认出）；")
    report.append("        纯字母数字的走词边界（Spindle 不中 Pin、Walnut 不中 Nut、")
    report.append("        ISOLATION 不中 ISO、Screwdriver 不中 Screw；但 GBT5782 能中 GBT）")

report.append("")
report.append("=== 自定义过程调用一致性（本项目前缀 U_/F_/X_/M_/G_/Cfg_，故无内置函数误报）===")
procs_def = set(re.findall(r"^(?:Private|Public|Friend)\s+(?:Sub|Function)\s+(\w+)",
                           codelines, re.M))
call_pat = r"\b((?:U_|F_|X_|M_|G_|Cfg_)\w+)\s*\(|\bCall\s+((?:U_|F_|X_|M_|G_|Cfg_)\w+)"
procs_call = set()
for a, b in re.findall(call_pat, codelines):
    procs_call.add(a or b)
procs_missing = sorted(procs_call - procs_def)
procs_unused = sorted(d for d in procs_def
                      if d not in procs_call and not d.startswith("SwAuto_"))
if procs_missing:
    for x in procs_missing:
        report.append("  *** 调用了但未定义: " + x)
else:
    report.append("  所有自定义过程调用都能找到定义 OK")
if procs_unused:
    report.append("  （定义了但没被调用: %s）" % ", ".join(procs_unused))

report.append("")
report.append("=== 危险 API 用法（这两个坑都实际踩过，写错会静默出错）===")
_danger = []

# ① Component2::IsHidden(ConsiderSuppressed) 的参数极容易写错：
#    传 True  → "隐藏 或 压缩 或 轻化" 都返回 True，轻化零件会被整批误杀
#    传 False → 只按可见性判断（正确用法，隐藏与压缩分开判）
for m in re.finditer(r"\.IsHidden\s*\(\s*([^)]*)\)", codelines):
    arg = m.group(1).strip()
    if arg.upper() != "FALSE":
        _danger.append("  *** IsHidden(%s) —— 必须传 False。"
                       "传 True 会把轻化零件一并误判成隐藏" % arg)

# ② Toolbox 判定：Component2::IsToolboxComponent 这个属性在 SW API 里根本不存在
if re.search(r"\.IsToolboxComponent\b", codelines):
    _danger.append("  *** 用了 IsToolboxComponent —— 该属性在 SolidWorks API 里不存在；"
                   "正确写法是 ModelDocExtension::ToolboxPartType")

# ③ 三个状态过滤开关必须都真的接进"要不要跳过"的判定函数，否则开关形同虚设。
#    注意检查点是 G_KindSkipped 而不是 G_CompBadState ——
#    后者只负责"读状态"（不看任何配置），前者才负责"按配置决定跳不跳"。
_kindskip = re.search(r"Private Function G_KindSkipped\b.*?\bEnd Function", codelines, re.S)
if not _kindskip:
    _danger.append("  *** 找不到 G_KindSkipped —— 装配体的隐藏/压缩/封套过滤可能被删了")
else:
    for _c in ["CFG_SKIP_HIDDEN", "CFG_SKIP_SUPPRESSED", "CFG_SKIP_ENVELOPE"]:
        if _c not in _kindskip.group(0):
            _danger.append("  *** G_KindSkipped 里没有引用 %s，这个开关不会生效" % _c)

# ④ 逐类勾选框的开关也得真的接进展开循环，否则勾选框永远不弹
_gather = re.search(r"Private Function G_ExpandAssembly\b.*?\bEnd Function", codelines, re.S)
if not _gather:
    _danger.append("  *** 找不到 G_ExpandAssembly —— 装配体展开逻辑可能被删了")
elif "CFG_ASK_SKIP_OPTIONS" not in _gather.group(0):
    _danger.append("  *** G_ExpandAssembly 里没有引用 CFG_ASK_SKIP_OPTIONS，勾选框不会弹")

# ⑤ 结果解析的判定顺序不能反。
#    「取消」回传的 C 只有一个字符，若先做"长度必须是 3"的校验，
#    它会被当成"没拿到结果"，函数返回 -2，用户按的取消就永远不生效。
#    这是个真实发生过的 bug（v2.2.0 到 v2.2.2），而且从界面上完全看不出来。
_ask = re.search(r"Private Function U_AskSkipOptions\b.*?\bEnd Function", codelines, re.S)
if not _ask:
    _danger.append("  *** 找不到 U_AskSkipOptions —— 装配体展开的勾选框可能被删了")
else:
    _askbody = _ask.group(0)
    _i_c = _askbody.find('sResp = "C"')
    _i_len = _askbody.find("Len(sResp)")
    if _i_c < 0:
        _danger.append('  *** U_AskSkipOptions 里找不到取消判断 sResp = "C"')
    elif _i_len >= 0 and _i_c > _i_len:
        _danger.append('  *** U_AskSkipOptions 里长度校验排在取消判断之前 —— '
                       'C 只有一个字符，会被误判成"没拿到结果"，取消永远不生效')
    for _w in (1, 2, 4):
        if ("U_AskSkipOptions = U_AskSkipOptions + %d" % _w) not in _askbody:
            _danger.append("  *** U_AskSkipOptions 里缺少权重 +%d（隐藏=1 压缩=2 封套=4）" % _w)

# ⑥ 跳过标记的开关和属性名都要真的接进判定，否则这个功能是哑的
_tagf = re.search(r"Private Function G_TagSkipped\b.*?\bEnd Function", codelines, re.S)
if not _tagf:
    _danger.append("  *** 找不到 G_TagSkipped —— 跳过标记功能可能被删了")
elif "CFG_USE_SKIP_TAG" not in _tagf.group(0):
    _danger.append("  *** G_TagSkipped 里没引用 CFG_USE_SKIP_TAG，总开关不会生效")
_tagd = re.search(r"Private Function G_DocHasTag\b.*?\bEnd Function", codelines, re.S)
if not _tagd:
    _danger.append("  *** 找不到 G_DocHasTag —— 读「跳过标记」属性那段可能被删了")
else:
    if "CFG_USE_SKIP_TAG" not in _tagd.group(0):
        _danger.append("  *** G_DocHasTag 里没引用 CFG_USE_SKIP_TAG")
    # ★ 两个页签都必须查。SolidWorks 属性对话框里「自定义」和「配置属性」是相邻两页，
    #   用户很容易把标记填在「配置属性」里；只查文件级的话，标记会形同虚设、
    #   零件照样被导出（v2.4.0 真实发生过的反馈）。
    #   注意这里查的是"组件引用的那个配置"（ReferencedConfiguration）而不是遍历全部 ——
    #   遍历太慢，一个零件可能几十个配置；全量遍历留给手动跑的自测宏。
    if "ReferencedConfiguration" not in _tagd.group(0):
        _danger.append('  *** G_DocHasTag 里没有 ReferencedConfiguration —— '
                       '只查了「自定义」页签，用户填在「配置属性」里就会认不出')

# 自测宏要把**所有配置**都翻一遍：它是手动跑的，慢一点无妨，
# 而且只有这样才诊断得出"标记填在别的配置上"这种情况。
_checktag = re.search(r"^Public Sub SwAuto_CheckTag\b.*?^End Sub", codelines, re.S | re.M)
if not _checktag:
    _danger.append("  *** 找不到自测宏 SwAuto_CheckTag")
elif "GetConfigurationNames" not in _checktag.group(0):
    _danger.append("  *** SwAuto_CheckTag 里没有 GetConfigurationNames —— "
                   '它应该把所有配置都翻一遍，否则诊断不出"标记填在别的配置里"')

# 属性名 / 值真正被用到的位置是 G_CpmHasTag（G_DocHasTag 只是调度）
_tagc = re.search(r"Private Function G_CpmHasTag\b.*?\bEnd Function", codelines, re.S)
if not _tagc:
    _danger.append("  *** 找不到 G_CpmHasTag —— 读属性的核心逻辑可能被删了")
else:
    # 注意用 \b 卡边界：CFG_SKIP_TAG 是 CFG_SKIP_TAG_VAL 的前缀，
    # 直接 in 判断会被后者蒙混过关
    for _c in ("CFG_SKIP_TAG", "CFG_SKIP_TAG_VAL"):
        if not re.search(r"\b%s\b" % _c, _tagc.group(0)):
            _danger.append("  *** G_CpmHasTag 里没引用 %s" % _c)

if _danger:
    report.extend(_danger)
else:
    report.append("  IsHidden 一律传 False（轻化零件不会被误杀）OK")
    report.append("  没有使用不存在的 IsToolboxComponent OK")
    report.append("  隐藏/压缩/封套 三个开关都接进了 G_KindSkipped OK")
    report.append("  勾选框的取消判断排在长度校验之前（C 不会被误判）OK")
    report.append("  掩码权重 隐藏=1 / 压缩=2 / 封套=4 齐全 OK")
    report.append("  跳过标记的开关与属性名都接进了判定 OK")

report.append("")
report.append("=== PowerShell 脚本（在 VB 里是拼字符串写的，只能还原出来验）===")


def _scan_vb_strings(line):
    """扫出一行 VB 代码里的字符串字面量（字符串内部的 "" 按转义处理）"""
    out, i, n = [], 0, len(line)
    while i < n:
        c = line[i]
        if c == "'":                       # 字符串外的单引号 = VB 注释开始
            break
        if c == '"':
            j, buf = i + 1, []
            while j < n:
                if line[j] == '"':
                    if j + 1 < n and line[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(line[j])
                j += 1
            out.append("".join(buf))
            i = j + 1
        else:
            i += 1
    return out


def _stub_runtime_exprs(line):
    """把 U_PsQuote(...) 这类**运行期才求值**的表达式换成一个 PS 字面量占位。

    这些值（标题、路径、个数）是执行时才拼进去的，静态还原拿不到。
    不换掉的话，还原出来会变成 `$Title = ` 这种右边空着的样子 —— 凭空多一个语法错误，
    检查就废了。括号要配对计数，因为里面还可能嵌着 CStr(IIf(...))。
    """
    key = "U_PsQuote("
    repl = '"' + "'X'" + '"'          # 在 VB 源码里就是 "'X'" 这么一个字符串字面量
    out, i = [], 0
    while True:
        j = line.find(key, i)
        if j < 0:
            out.append(line[i:])
            break
        out.append(line[i:j])
        k, depth = j + len(key), 1
        while k < len(line) and depth > 0:
            if line[k] == "(":
                depth += 1
            elif line[k] == ")":
                depth -= 1
            k += 1
        out.append(repl)
        i = k
    return "".join(out)


def _vb_func_text(fname):
    """把某个 VB 函数里 s = s & "..." & vbCrLf 的拼接还原成完整文本"""
    m = re.search(r"^Private Function %s\b.*?^End Function" % fname, src, re.S | re.M)
    if not m:
        return None
    parts = []
    for ln in m.group(0).split("\n"):
        segs = _scan_vb_strings(_stub_runtime_exprs(ln))
        if not segs:
            continue
        parts.append("".join(segs))
        if "vbCrLf" in ln:
            parts.append("\n")
    return "".join(parts)


def _check_ps(fname, label):
    import os as _os
    import tempfile
    import subprocess

    ps = _vb_func_text(fname)
    if ps is None:
        report.append("  *** 找不到 %s" % fname)
        return
    bad = []
    # 参数必须**内嵌在脚本正文里**。一旦出现 param(，说明有人改回"命令行传参"了 ——
    # 那条路在 cmd 的单引号解析上真翻过车：WScript.Shell.Run 经 cmd 启动 powershell，
    # cmd 不认单引号，参数值会带着引号进脚本，结果文件写到别处，
    # 表现为"新式对话框弹了又弹一次旧式的"。详见 src/25_dialog.bas 文件头部的说明。
    if "param(" in ps:
        bad.append("脚本里又出现了 param( —— 参数必须内嵌，命令行传参会踩 cmd 单引号的坑")
    # here-string 是可选的（纯 WinForms 那段就没有），有的话必须成对且最多一组
    _no, _nc = ps.count("@'"), ps.count("'@")
    if _no != _nc:
        bad.append("here-string 标记 @' / '@ 数量不匹配（%d vs %d）" % (_no, _nc))
    elif _no > 1:
        bad.append("here-string 标记出现 %d 组，预期最多 1 组" % _no)
    # '@ 必须顶格在行首 —— 这是 PowerShell 的硬要求，缩进了就是语法错误
    if _no > 0 and not any(l == "'@" for l in ps.split("\n")):
        bad.append("here-string 结束标记 '@ 没有顶格")
    for a, b, nm in [("{", "}", "大括号"), ("(", ")", "圆括号")]:
        if ps.count(a) != ps.count(b):
            bad.append("%s不配平（%d vs %d）" % (nm, ps.count(a), ps.count(b)))
    if bad:
        for x in bad:
            report.append("  *** %s：%s" % (label, x))
        return

    # 静态检查过了，再交给 PowerShell 自己的解析器验一遍语法
    try:
        tmp = _os.path.join(tempfile.gettempdir(), "swauto_syntax_%s.ps1" % fname)
        with open(tmp, "w", encoding="utf-8-sig") as fh:
            fh.write(ps)
        q = tmp.replace("\\", "\\\\")
        expr = ("$e=$null;[void][System.Management.Automation.Language.Parser]::ParseFile("
                "'%s',[ref]$null,[ref]$e);if($e){$e|ForEach-Object{$_.Message}}else{'PARSE-OK'}" % q)
        r = subprocess.run(["powershell.exe", "-NoProfile", "-Command", expr],
                           capture_output=True, timeout=40)
        out = r.stdout.decode("utf-8", "replace").strip()
        if not out:
            out = r.stdout.decode("gbk", "replace").strip()
        if "PARSE-OK" in out:
            report.append("  %s：here-string 与括号配平 OK；PowerShell 解析通过 OK（%d 字符）"
                          % (label, len(ps)))
        else:
            report.append("  *** %s：PowerShell 报语法错误 → %s" % (label, out[:220]))
    except Exception as e:
        report.append("  %s：静态检查 OK（本机没跑成 PowerShell 解析：%s）" % (label, e))


_check_ps("U_FolderPickerScript", "文件夹选择器")
_check_ps("U_AskSkipScript", "展开选项勾选框")

report.append("")
report.append("=== 入口宏（必须 Public Sub 且无参数）===")
for idx, ln in enumerate(lines, 1):
    if ln.startswith("Public Sub"):
        report.append("  行 %4d: %s" % (idx, ln.strip()))

# 每个"干活的"入口都必须在开头设置 m_ModeTag —— 所有弹窗标题靠它区分模式，
# 这样"快捷键绑到了别的入口"能一眼看出来（实际发生过：绑成文件夹模式还以为坏了）。
report.append("")
report.append("=== 入口宏的模式标识（漏了 = 弹窗标题分不清是哪个模式）===")
# SwAuto_Help（纯说明）和 SwAuto_CheckTag（自测工具）都不参与导出，不需要模式标识
EXEMPT_ENTRIES = {"SwAuto_Help", "SwAuto_CheckTag"}
_mode_bad = 0
for m in re.finditer(r"^Public Sub\s+(SwAuto_\w+)\s*\(\s*\)", codelines, re.M):
    name = m.group(1)
    tail = codelines[m.end():]
    nxt = re.search(r"^(?:Private|Public|Friend)\s+(?:Sub|Function)\s+(\w+)", tail, re.M)
    body = tail[:nxt.start()] if nxt else tail
    tag = re.search(r'm_ModeTag\s*=\s*"([^"]*)"', body)
    if tag:
        report.append("  OK %-26s -> %s" % (name, tag.group(1)))
    elif name in EXEMPT_ENTRIES:
        report.append("  -- %-26s （说明宏，无需模式标识）" % name)
    else:
        report.append("  *** %s 没有设置 m_ModeTag" % name)
        _mode_bad += 1
if _mode_bad == 0:
    report.append("  弹窗标题都能标明模式 OK")

out = "\n".join(report)
open("dist/_check_report.txt", "w", encoding="utf-8").write(out)
print(out)
