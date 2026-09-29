import glob, os, sys

# ---------------------------------------------------------------------------
# VBE「文件 > 导入文件」读取 .bas 时，是按【系统 ANSI 代码页】解码的。
# 简体中文 Windows 下就是 cp936(GBK)。
#
# 这里千万不能写 UTF-8：
#   1) 中文注释和字符串会全部乱码；
#   2) 更致命的是——中文标点（如「。」的 UTF-8 字节 E3 80 82）在 GBK 里，
#      收尾字节 82/83/8x 属于 GBK 前导字节，会跟紧随其后的 ASCII 字符
#      （通常就是字符串的结束引号 " 或连接符 &）合成一个汉字，把引号吃掉，
#      字符串失去结尾 → VBE 报「编译错误：语法错误」。
# 所以：GBK 编码 + CRLF 行尾。
# ---------------------------------------------------------------------------
ENC = "gbk"
OUT = "dist/SWautoExport.bas"

os.makedirs("dist", exist_ok=True)
files = sorted(glob.glob("src/*.bas"))
if not files:
    sys.exit("ERROR: 没找到 src/*.bas")

parts = []
for f in files:
    d = open(f, "rb").read().replace(b"\r\n", b"\n").rstrip(b"\n")
    parts.append(d)

data = b"\n\n".join(parts) + b"\n"
data = data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
text = data.decode("utf-8")

# --- 检查是否有 GBK 编不出的字符，提前暴露，而不是静默丢字 ---
bad = []
for ch in set(text):
    try:
        ch.encode(ENC)
    except UnicodeEncodeError:
        bad.append(ch)
if bad:
    print("!! 这些字符 GBK 编不出来，会丢字：" + " ".join("U+%04X" % ord(c) for c in sorted(bad)))

enc = text.encode(ENC, "replace")
open(OUT, "wb").write(enc)

# --- 备用通道：UTF-8 with BOM 的纯文本，用于「打开→全选→复制→粘贴到 VBE」---
# 剪贴板是 Unicode，完全不经过代码页，所以这条通道 100% 不受系统区域设置影响。
PASTE = "dist/SWautoExport_粘贴用.txt"
open(PASTE, "wb").write(b"\xef\xbb\xbf" + text.replace("\r\n", "\n").encode("utf-8"))

# --- 回读校验 ---
back = open(OUT, "rb").read().decode(ENC)
lines = back.split("\r\n")
report = []
report.append("output : %s  (%s, CRLF)" % (OUT, ENC.upper()))
report.append("files  : " + ", ".join(files))
report.append("lines  : %d" % (len(lines) - 1))
report.append("bytes  : %d" % len(enc))
report.append("round-trip decode OK: %s" % ("YES" if back == text else "NO"))
report.append("Option Explicit 次数: %d" % text.count("Option Explicit"))

procs = []
for ln in lines:
    s = ln.strip()
    for kw in ("Private Sub ", "Public Sub ", "Private Function ", "Public Function "):
        if s.startswith(kw):
            procs.append(s.split("(")[0].split()[-1])
dup = sorted({p for p in procs if procs.count(p) > 1})
report.append("过程数: %d   重名: %s" % (len(procs), ", ".join(dup) if dup else "无"))

out = "\n".join(report)
open("dist/_build_report.txt", "w", encoding="utf-8").write(out)
print(out)
