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
report.append("=== 类型库依赖 ===")
for x in sorted(set(re.findall(r"\b(sw[A-Za-z]+_e)\b", src))):
    report.append("  " + x)
report.append("--- 枚举成员 ---")
for x in sorted(set(re.findall(r"\bsw[A-Za-z]+_e\.(sw[A-Za-z0-9_]+)", src))):
    report.append("  " + x)

report.append("")
report.append("=== 入口宏（必须 Public Sub 且无参数）===")
for idx, ln in enumerate(lines, 1):
    if ln.startswith("Public Sub"):
        report.append("  行 %4d: %s" % (idx, ln.strip()))

out = "\n".join(report)
open("dist/_check_report.txt", "w", encoding="utf-8").write(out)
print(out)
