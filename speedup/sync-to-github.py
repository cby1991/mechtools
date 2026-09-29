#!/usr/bin/env python
"""把本地 speedup 仓库同步到 mechtools 仓库的 speedup/ 子目录。

为什么需要这个脚本
------------------
git 本身不支持「把 A 仓库的根，推到 B 仓库的某个子目录」。本项目就是这种情况：

    本地  D:\\WorkSpace\\projectmanage\\speedup   ← 仓库根 == speedup 的内容
    远端  cby1991/mechtools
            ├── LICENSE
            ├── README.md
            └── speedup/                            ← 我们的内容要落在这里

`git subtree push --prefix=...` 要求 prefix 在**当前仓库里真实存在**，
但本地根本身就是 speedup，没有可指的 `speedup/` 子目录，所以用不了。

这个脚本用「外壳仓库」的办法绕过去：

    1. 在临时目录建一个空仓库，接上远端 mechtools
    2. 检出远端 main
    3. 用 `git bundle` 把本地仓库的对象搬进去（纯对象传输，不碰工作区）
    4. `git read-tree --prefix=speedup/ -u <本地HEAD>` 把内容挂到子目录
    5. 提交 + push

本地仓库**完全不受影响**：不建分支、不改工作区、不加远端引用之外的东西。

用法
----
    uv run python sync-to-github.py                # 正常同步
    uv run python sync-to-github.py --dry-run      # 只做到提交，不推送
    uv run python sync-to-github.py --message "..."  # 自定义提交信息

硬约束
------
- R23：所有临时文件写到 %TEMP%，绝不落在仓库里
- 不吞 stdout：每一步的真实输出都打出来，退出码非 0 就停
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# --- 配置（改这里就能换仓库/子目录）-----------------------------------------

REMOTE_URL = "https://github.com/cby1991/mechtools.git"
BRANCH = "main"
PREFIX = "speedup"          # 我们要落到远端的子目录名
COMMITTER_NAME = "cby"
COMMITTER_EMAIL = "boyang_chen@163.com"

# 沙箱会注入 http_proxy，导致 git 网络命令 502。跑 git 时剥离掉。
_PROXY_VARS = ("http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY")

DEFAULT_MESSAGE = "Sync speedup from local working copy"


def run(
    *args: str,
    cwd: Path | None = None,
    check: bool = True,
    capture: bool = False,
) -> subprocess.CompletedProcess[str]:
    """跑一条外部命令，默认剥离代理环境变量。"""
    env = {k: v for k, v in os.environ.items() if k not in _PROXY_VARS}
    kwargs: dict = {"cwd": cwd, "env": env, "text": True}
    if capture:
        kwargs["capture_output"] = True
    result = subprocess.run(args, **kwargs)
    if check and result.returncode != 0:
        sys.exit(f"\n✗ 命令失败（退出码 {result.returncode}）：{' '.join(args)}")
    return result


def git(*args: str, cwd: Path, capture: bool = False) -> str:
    """跑一条 git 命令，返回 stdout（capture=True 时）。"""
    result = run("git", *args, cwd=cwd, capture=capture)
    return (result.stdout or "").strip() if capture else ""


def step(n: int, text: str) -> None:
    print(f"\n[{n}] {text}", flush=True)


def info(text: str) -> None:
    """打一行进度。必须 flush —— 子进程的 stdout 是行缓冲/直接写 fd，
    而我们自己的 stdout 被管道重定向时是全缓冲，不 flush 就会导致
    git 的输出跑到我们的日志前面，看着像错乱。"""
    print(text, flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dry-run", action="store_true",
                        help="只做到提交，不推送")
    parser.add_argument("--message", "-m", default=None,
                        help="提交信息（默认自动生成）")
    parser.add_argument("--remote", default=REMOTE_URL, help="远端仓库地址")
    parser.add_argument("--prefix", default=PREFIX, help="远端子目录名")
    args = parser.parse_args()

    repo = Path.cwd()
    if not (repo / ".git").exists():
        sys.exit(f"✗ 当前目录不是 git 仓库：{repo}")

    # --- 0. 本地必须是干净的，否则同步上去的内容和本地不一致 ----------------

    step(0, "检查本地工作区")
    dirty = git("status", "--porcelain", cwd=repo, capture=True)
    if dirty:
        info("本地有未提交的改动：")
        info(dirty)
        sys.exit("✗ 请先提交（或 stash），再同步。否则推上去的不是你本地看到的样子。")
    head = git("rev-parse", "HEAD", cwd=repo, capture=True)
    info(f"  本地 HEAD = {head[:12]}")

    # --- 1. 准备临时目录 -----------------------------------------------------

    tmp_root = Path(tempfile.mkdtemp(prefix="mechtools-sync-"))
    shell = tmp_root / "shell"
    bundle = tmp_root / "speedup.bundle"
    info(f"  临时目录 = {tmp_root}")

    try:
        # --- 2. 建外壳仓库并接上远端 ----------------------------------------

        step(1, "建外壳仓库")
        shell.mkdir()
        git("init", "-q", "-b", BRANCH, ".", cwd=shell)
        git("config", "user.name", COMMITTER_NAME, cwd=shell)
        git("config", "user.email", COMMITTER_EMAIL, cwd=shell)
        git("remote", "add", "origin", args.remote, cwd=shell)
        info(f"  {args.remote}")

        step(2, "拉取远端")
        git("fetch", "origin", BRANCH, "--depth=1", cwd=shell)
        git("checkout", "-q", "-b", BRANCH, f"origin/{BRANCH}", cwd=shell)
        info("  已检出远端分支")

        # --- 3. 把本地对象搬进来（bundle 是纯对象传输）-----------------------

        step(3, "导入本地对象")
        git("bundle", "create", str(bundle), "HEAD", cwd=repo)
        git("fetch", str(bundle), "HEAD", cwd=shell)
        info("  已导入")

        # --- 4. 把内容挂到子目录 ---------------------------------------------
        #
        # 两种情形：
        #   a) 远端还没有 <prefix>/    → read-tree --prefix 直接挂上
        #   b) 远端已经有 <prefix>/    → 必须先清空索引里那块，否则 read-tree
        #      报 "Entry 'speedup/x' overlaps with 'speedup/x'. Cannot bind."
        #
        # 情形 b 就是「更新已有子目录」，是日常同步的常态，不是异常。

        step(4, f"挂载到 {args.prefix}/")
        existing = git("ls-files", f"{args.prefix}/", cwd=shell, capture=True)
        if existing:
            n_old = existing.count("\n") + 1
            info(f"  远端已有 {args.prefix}/（{n_old} 个文件），先清空索引")
            git("rm", "-r", "--cached", "-q", "--ignore-unmatch", f"{args.prefix}/",
                cwd=shell)
            # 工作区里那些文件也要清掉，否则 read-tree -u 会被残留文件挡住
            stale = shell / args.prefix
            if stale.exists():
                shutil.rmtree(stale)
        git("read-tree", f"--prefix={args.prefix}/", "-u", head, cwd=shell)
        count = git("ls-files", cwd=shell, capture=True).count("\n") + 1
        info(f"  挂载完成，索引里共 {count} 个文件")

        # --- 5. 提交 ---------------------------------------------------------

        step(5, "提交")
        git("add", "-A", cwd=shell)
        message = args.message or DEFAULT_MESSAGE
        msg_file = tmp_root / "commitmsg.txt"
        msg_file.write_text(message + "\n", encoding="utf-8")
        git("commit", "-F", str(msg_file), cwd=shell)
        author = git("log", "-1", "--format=%an <%ae>", cwd=shell, capture=True)
        info(f"  提交者：{author}")

        if args.dry_run:
            step(6, "dry-run：跳过推送")
            info(f"  外壳仓库留在：{shell}")
            return 0

        # --- 6. 推送 ---------------------------------------------------------

        step(6, "推送")
        git("push", "origin", BRANCH, cwd=shell)

        step(7, "完成")
        info(f"  https://github.com/cby1991/mechtools/tree/main/{args.prefix}")
        return 0

    finally:
        # 成功就清理；dry-run 时保留 tmp_root 供人工查看
        if tmp_root.exists() and not args.dry_run:
            shutil.rmtree(tmp_root, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
