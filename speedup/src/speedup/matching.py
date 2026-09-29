"""补料清单名称 ↔ 磁盘文件 的匹配与拷贝规划。

分层设计（为了可测）
--------------------
* :func:`match_score` / :func:`find_best` / :func:`build_matches` —— 纯函数，不碰磁盘写入。
* :func:`scan_files` —— 只读磁盘。
* :func:`plan_copies` —— **纯函数**，只决定「该拷什么、该跳过什么」，不真的拷。
* :func:`execute_copies` —— 唯一真正写磁盘的一层。

把「决策」和「执行」拆开的好处：`--dry-run` 只需要调用 :func:`plan_copies`，
测试也可以在临时目录里精确断言每一步的结果。

匹配强度（``min_score``）
-------------------------
===  ============================================================
 4   只认文件名与名称**完全一致**（归一化后）
 3   另外允许「名称 + 版本号后缀」，如 ``驱动连杆`` → ``驱动连杆V1.1.STEP``（默认）
 1   宽松：互相包含也算，结果会标 ``!`` 提示人工核对
===  ============================================================
"""

from __future__ import annotations

import os
import re
import shutil
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from typing import Final

from .config import ALL_SCAN_EXTS, DEFAULT_MIN_SCORE, DRAW_EXTS, MODEL_EXTS, SKIP_PREFIXES

__all__ = [
    "LEVEL_OK",
    "LEVEL_FUZZY",
    "LEVEL_PARTIAL",
    "LEVEL_MISS",
    "LEVEL_MARKERS",
    "ScannedFile",
    "MatchResult",
    "CopyAction",
    "CopyReport",
    "normalize",
    "match_score",
    "find_best",
    "scan_files",
    "build_matches",
    "plan_copies",
    "execute_copies",
    "find_shared_matches",
    "unused_files",
]

#: 匹配等级，供 :func:`speedup.console.render_table` 标注用
LEVEL_OK: Final[str] = "ok"
LEVEL_FUZZY: Final[str] = "fuzzy"
LEVEL_PARTIAL: Final[str] = "partial"
LEVEL_MISS: Final[str] = "miss"

#: 每个等级在表格行首打的标记
LEVEL_MARKERS: Final[Mapping[str, str]] = {
    LEVEL_OK: "  ",
    LEVEL_FUZZY: "! ",
    LEVEL_PARTIAL: "! ",
    LEVEL_MISS: "X ",
}

_WHITESPACE = re.compile(r"[\s\u3000_]")


@dataclass(frozen=True, slots=True)
class ScannedFile:
    """扫描到的候选文件。"""

    path: str
    stem: str
    ext: str  # 小写且含点，例如 ".step"

    @property
    def is_model(self) -> bool:
        return self.ext in MODEL_EXTS

    @property
    def is_drawing(self) -> bool:
        return self.ext in DRAW_EXTS


@dataclass(frozen=True, slots=True)
class MatchResult:
    """清单里一个名称的匹配结果。"""

    name: str
    model: str | None = None
    draw: str | None = None
    model_candidates: int = 0
    draw_candidates: int = 0
    score: int = 0
    status: str = ""
    level: str = LEVEL_MISS

    @property
    def has_model(self) -> bool:
        return bool(self.model)

    @property
    def has_draw(self) -> bool:
        return bool(self.draw)

    @property
    def complete(self) -> bool:
        return self.has_model and self.has_draw


@dataclass(frozen=True, slots=True)
class CopyAction:
    """一条拷贝决策。``kind`` 取 ``copy`` / ``skip_same_name`` / ``inplace``。"""

    name: str
    src: str
    dst: str
    kind: str


@dataclass
class CopyReport:
    """一次拷贝执行的统计结果。"""

    copied: list[tuple[str, str]] = field(default_factory=list)
    skipped: list[tuple[str, str]] = field(default_factory=list)
    inplace: list[tuple[str, str]] = field(default_factory=list)
    failed: list[tuple[str, str, str]] = field(default_factory=list)

    @property
    def total_attempted(self) -> int:
        return len(self.copied) + len(self.skipped) + len(self.inplace) + len(self.failed)


def normalize(text: object) -> str:
    """归一化：转小写、去掉所有空白（含全角空格）和下划线。"""
    return _WHITESPACE.sub("", str(text).lower())


def match_score(name: str, stem: str) -> int:
    """给「清单名称 ↔ 文件名主干」打分。见模块 docstring 的强度表。"""
    left, right = normalize(name), normalize(stem)
    if not left or not right:
        return 0
    if left == right:
        return 4
    if len(left) >= 2 and right.startswith(left):
        return 3
    if len(right) >= 2 and left.startswith(right):
        return 2
    if len(left) >= 2 and len(right) >= 2 and (left in right or right in left):
        return 1
    return 0


def find_best(
    name: str,
    candidates: Sequence[ScannedFile],
    min_score: int = DEFAULT_MIN_SCORE,
) -> tuple[str | None, int]:
    """在候选里挑最佳匹配。

    同分时取 **修改时间更新** 的那个（同一零件存在多个版本时，通常新的才对）。

    :returns: ``(最佳路径或 None, 达到 min_score 的候选个数)``
    """
    scored: list[tuple[int, float, str]] = []
    for candidate in candidates:
        score = match_score(name, candidate.stem)
        if score < min_score:
            continue
        try:
            mtime = os.path.getmtime(candidate.path)
        except OSError:
            mtime = 0.0
        scored.append((score, mtime, candidate.path))
    if not scored:
        return None, 0
    scored.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return scored[0][2], len(scored)


def _score_of(name: str, path: str | None) -> int:
    if not path:
        return 0
    return match_score(name, os.path.splitext(os.path.basename(path))[0])


def scan_files(
    root: str | os.PathLike[str],
    recursive: bool = True,
    exts: Iterable[str] = ALL_SCAN_EXTS,
) -> list[ScannedFile]:
    """扫描文件夹里的模型与图纸，跳过 ``~$`` / ``.`` 开头的文件。"""
    wanted = {e.lower() for e in exts}
    found: list[ScannedFile] = []

    if recursive:
        walker: Iterable[tuple[str, list[str], list[str]]] = os.walk(root)
    else:
        walker = [(str(root), [], [n for n in os.listdir(root) if os.path.isfile(os.path.join(root, n))])]

    for dirpath, _dirnames, filenames in walker:
        for filename in filenames:
            if filename.startswith(SKIP_PREFIXES):
                continue
            stem, ext = os.path.splitext(filename)
            ext_lower = ext.lower()
            if ext_lower in wanted:
                found.append(ScannedFile(path=os.path.join(dirpath, filename), stem=stem, ext=ext_lower))
    return found


def build_matches(
    names: Sequence[str],
    files: Sequence[ScannedFile],
    min_score: int = DEFAULT_MIN_SCORE,
) -> list[MatchResult]:
    """为每个名称找模型与图纸，标记匹配等级。"""
    models = [f for f in files if f.is_model]
    drawings = [f for f in files if f.is_drawing]

    results: list[MatchResult] = []
    for name in names:
        # 先用最宽松的强度算一遍，以便区分「压根没有相似文件」和「有相似的但不敢匹配」
        loose_model, model_candidates = find_best(name, models, min_score=1)
        loose_draw, draw_candidates = find_best(name, drawings, min_score=1)

        model = loose_model if _score_of(name, loose_model) >= min_score else None
        draw = loose_draw if _score_of(name, loose_draw) >= min_score else None

        score = max(_score_of(name, model), _score_of(name, draw))

        if not model and not draw:
            status = "未找到(有相似文件)" if (loose_model or loose_draw) else "未找到"
            level = LEVEL_MISS
        elif not model or not draw:
            status, level = "只找到一半", LEVEL_PARTIAL
        elif score >= 4:
            status, level = "完全匹配", LEVEL_OK
        elif score == 3:
            status, level = "匹配(带版本号)", LEVEL_OK
        else:
            status, level = "模糊匹配(请核对)", LEVEL_FUZZY

        results.append(
            MatchResult(
                name=name,
                model=model,
                draw=draw,
                model_candidates=model_candidates,
                draw_candidates=draw_candidates,
                score=score,
                status=status,
                level=level,
            )
        )
    return results


def _same_path(left: str, right: str) -> bool:
    return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right))


def plan_copies(
    results: Sequence[MatchResult],
    target_dir: str | os.PathLike[str],
) -> list[CopyAction]:
    """决定每个命中文件该怎么处理：**不写磁盘**。

    规则的顺序很重要：

    1. 源和目标是同一个文件 → ``inplace``（本来就在目标文件夹，跳过）
    2. 目标已存在同名文件 → ``skip_same_name``（**不覆盖**，防止冲掉人工改过的表）
    3. 其余 → ``copy``
    """
    target = os.fspath(target_dir)
    actions: list[CopyAction] = []
    for result in results:
        for src in (result.model, result.draw):
            if not src:
                continue
            dst = os.path.join(target, os.path.basename(src))
            if _same_path(src, dst):
                kind = "inplace"
            elif os.path.exists(dst):
                kind = "skip_same_name"
            else:
                kind = "copy"
            actions.append(CopyAction(name=result.name, src=src, dst=dst, kind=kind))
    return actions


def execute_copies(actions: Sequence[CopyAction]) -> CopyReport:
    """执行 :func:`plan_copies` 产出的决策。只有 ``kind == "copy"`` 会真的写磁盘。"""
    report = CopyReport()
    for action in actions:
        filename = os.path.basename(action.dst)
        if action.kind == "inplace":
            report.inplace.append((action.name, filename))
        elif action.kind == "skip_same_name":
            report.skipped.append((action.name, filename))
        else:
            try:
                shutil.copy2(action.src, action.dst)
                report.copied.append((action.name, filename))
            except Exception as exc:  # 权限 / 占用 / 磁盘满
                report.failed.append((action.name, filename, str(exc)))
    return report


def find_shared_matches(results: Sequence[MatchResult]) -> list[tuple[str, str, str]]:
    """找出「被多个不同名称同时匹配到」的文件，用于提醒可能拷错。

    :returns: ``[(名称A, 名称B, 文件名), ...]``
    """
    owner: dict[str, str] = {}
    clashes: list[tuple[str, str, str]] = []
    for result in results:
        for path in (result.model, result.draw):
            if not path:
                continue
            key = os.path.normcase(os.path.abspath(path))
            previous = owner.get(key)
            if previous is not None and previous != result.name:
                clashes.append((previous, result.name, os.path.basename(path)))
            else:
                owner[key] = result.name
    return clashes


def unused_files(
    files: Sequence[ScannedFile],
    results: Sequence[MatchResult],
) -> list[ScannedFile]:
    """扫描到但没有任何名称匹配上、因此不会被拷贝的文件。"""
    used = {
        os.path.normcase(os.path.abspath(path))
        for result in results
        for path in (result.model, result.draw)
        if path
    }
    return [f for f in files if os.path.normcase(os.path.abspath(f.path)) not in used]
