"""匹配与拷贝规划测试。

这些是补料工具最容易出错的地方（拷错文件 / 漏拷 / 覆盖人工修改），
所以覆盖得细一些。
"""

from __future__ import annotations

import os

import pytest

from speedup.matching import (
    LEVEL_FUZZY,
    LEVEL_MISS,
    LEVEL_OK,
    LEVEL_PARTIAL,
    CopyAction,
    MatchResult,
    ScannedFile,
    build_matches,
    execute_copies,
    find_best,
    find_shared_matches,
    match_score,
    normalize,
    plan_copies,
    scan_files,
    unused_files,
)


class TestNormalize:
    def test_lowercases(self):
        assert normalize("ABC") == "abc"

    def test_removes_spaces_and_underscores(self):
        assert normalize("腕 关_节") == "腕关节"

    def test_removes_fullwidth_space(self):
        assert normalize("腕\u3000关节") == "腕关节"

    def test_handles_non_string(self):
        assert normalize(123) == "123"


class TestMatchScore:
    def test_identical_is_four(self):
        assert match_score("垫片", "垫片") == 4

    def test_case_and_space_insensitive(self):
        assert match_score("head front", "Head_Front") == 4

    def test_name_plus_version_suffix_is_three(self):
        assert match_score("驱动连杆", "驱动连杆V1.1") == 3

    def test_name_longer_than_stem_is_two(self):
        assert match_score("垫片A", "垫片") == 2

    def test_containment_is_one(self):
        """「基座」在「末端基座」里出现，算最低一档匹配（需人工核对）。"""
        assert match_score("基座", "末端基座") == 1

    def test_unrelated_is_zero(self):
        assert match_score("垫片", "法兰盘") == 0

    def test_single_char_name_never_fuzzy_matches(self):
        """单个字符太容易误匹配，只允许完全相同。"""
        assert match_score("A", "AB") == 0

    def test_empty_inputs(self):
        assert match_score("", "垫片") == 0
        assert match_score("垫片", "") == 0


class TestScanFiles:
    def test_scans_recursively(self, tmp_path, make_files):
        make_files(tmp_path, ["a.STEP", "b.pdf", "c.txt"])
        make_files(tmp_path / "sub", ["d.stp"])
        found = scan_files(tmp_path)
        names = sorted(f.stem for f in found)
        assert names == ["a", "b", "d"]

    def test_non_recursive(self, tmp_path, make_files):
        make_files(tmp_path, ["a.STEP"])
        make_files(tmp_path / "sub", ["d.stp"])
        assert len(scan_files(tmp_path, recursive=False)) == 1

    def test_skips_excel_temp_and_hidden_files(self, tmp_path, make_files):
        make_files(tmp_path, ["~$a.STEP", ".hidden.STEP", "real.STEP"])
        found = scan_files(tmp_path)
        assert [f.stem for f in found] == ["real"]

    def test_extension_is_lowercased(self, tmp_path, make_files):
        make_files(tmp_path, ["A.STEP"])
        assert scan_files(tmp_path)[0].ext == ".step"

    def test_only_requested_extensions(self, tmp_path, make_files):
        make_files(tmp_path, ["a.step", "b.pdf"])
        assert len(scan_files(tmp_path, exts=(".pdf",))) == 1

    def test_empty_folder(self, tmp_path):
        assert scan_files(tmp_path) == []

    def test_classifies_models_and_drawings(self, tmp_path, make_files):
        make_files(tmp_path, ["a.step", "b.pdf"])
        found = {f.ext: f for f in scan_files(tmp_path)}
        assert found[".step"].is_model and not found[".step"].is_drawing
        assert found[".pdf"].is_drawing and not found[".pdf"].is_model


class TestFindBest:
    def test_picks_highest_score(self, scanned):
        candidates = scanned(["驱动连杆V1.1.step", "驱动连杆.step"])
        best, count = find_best("驱动连杆", candidates, min_score=3)
        assert best is not None
        assert os.path.basename(best) == "驱动连杆.step"
        assert count == 2

    def test_respects_min_score(self, scanned):
        candidates = scanned(["驱动连杆V1.1.step"])
        assert find_best("驱动连杆", candidates, min_score=4)[0] is None
        assert find_best("驱动连杆", candidates, min_score=3)[0] is not None

    def test_returns_none_and_zero_when_nothing_matches(self, scanned):
        assert find_best("不存在", scanned(["a.step"])) == (None, 0)

    def test_ties_prefer_newer_mtime(self, tmp_path, make_files):
        (older,) = make_files(tmp_path, ["垫片A.step"])
        (newer,) = make_files(tmp_path, ["垫片B.step"])
        os.utime(older, (1_000_000, 1_000_000))
        os.utime(newer, (2_000_000, 2_000_000))
        candidates = [
            ScannedFile(path=str(older), stem="垫片A", ext=".step"),
            ScannedFile(path=str(newer), stem="垫片B", ext=".step"),
        ]
        best, count = find_best("垫片", candidates, min_score=1)
        assert best == str(newer)
        assert count == 2


class TestBuildMatches:
    def test_complete_match(self, scanned):
        files = scanned(["关节支撑板.STEP", "关节支撑板.pdf"])
        (result,) = build_matches(["关节支撑板"], files)
        assert result.level == LEVEL_OK
        assert result.status == "完全匹配"
        assert result.complete

    def test_version_suffix_counts_as_ok(self, scanned):
        files = scanned(["驱动连杆V1.1.STEP", "驱动连杆V1.1.pdf"])
        (result,) = build_matches(["驱动连杆"], files)
        assert result.level == LEVEL_OK
        assert result.status == "匹配(带版本号)"
        assert result.complete

    def test_partial_when_drawing_missing(self, scanned):
        (result,) = build_matches(["垫片"], scanned(["垫片.STEP"]))
        assert result.level == LEVEL_PARTIAL
        assert result.status == "只找到一半"
        assert result.has_model and not result.has_draw

    def test_miss(self, scanned):
        (result,) = build_matches(["不存在的零件"], scanned(["垫片.STEP"]))
        assert result.level == LEVEL_MISS
        assert result.status == "未找到"

    def test_miss_with_similar_files_is_flagged_differently(self, scanned):
        """有相似文件但强度不够时，要提示「有相似文件」，方便人工去核。"""
        (result,) = build_matches(["基座"], scanned(["末端基座.step"]), min_score=3)
        assert result.level == LEVEL_MISS
        assert result.status == "未找到(有相似文件)"

    def test_loose_strength_marks_fuzzy(self, scanned):
        files = scanned(["末端基座.step", "末端基座.pdf"])
        (result,) = build_matches(["基座"], files, min_score=1)
        assert result.level == LEVEL_FUZZY
        assert result.status == "模糊匹配(请核对)"
        assert result.complete  # 模型和图纸都找到了，只是名字对不严

    def test_preserves_name_order(self, scanned):
        files = scanned(["a.step", "b.step"])
        results = build_matches(["b", "a"], files)
        assert [r.name for r in results] == ["b", "a"]

    def test_records_candidate_counts(self, scanned):
        files = scanned(["垫片V1.step", "垫片V2.step"])
        (result,) = build_matches(["垫片"], files)
        assert result.model_candidates == 2


class TestPlanCopies:
    def test_plans_copy_when_target_missing(self, tmp_path, make_files):
        (source,) = make_files(tmp_path / "src", ["垫片.STEP"])
        target = tmp_path / "dst"
        target.mkdir()
        result = MatchResult(name="垫片", model=str(source))

        (action,) = plan_copies([result], target)
        assert action.kind == "copy"
        assert action.dst == str(target / "垫片.STEP")

    def test_skips_existing_same_name(self, tmp_path, make_files):
        (source,) = make_files(tmp_path / "src", ["垫片.STEP"])
        target = tmp_path / "dst"
        make_files(target, ["垫片.STEP"])

        (action,) = plan_copies([MatchResult(name="垫片", model=str(source))], target)
        assert action.kind == "skip_same_name"

    def test_inplace_when_source_is_target(self, tmp_path, make_files):
        (source,) = make_files(tmp_path, ["垫片.STEP"])
        (action,) = plan_copies([MatchResult(name="垫片", model=str(source))], tmp_path)
        assert action.kind == "inplace"

    def test_handles_both_model_and_draw(self, tmp_path, make_files):
        (model,) = make_files(tmp_path / "src", ["垫片.STEP"])
        (draw,) = make_files(tmp_path / "src", ["垫片.pdf"])
        target = tmp_path / "dst"
        target.mkdir()
        actions = plan_copies([MatchResult(name="垫片", model=str(model), draw=str(draw))], target)
        assert len(actions) == 2

    def test_ignores_missing_paths(self, tmp_path):
        assert plan_copies([MatchResult(name="垫片")], tmp_path) == []

    def test_does_not_write_anything(self, tmp_path, make_files):
        (source,) = make_files(tmp_path / "src", ["垫片.STEP"])
        target = tmp_path / "dst"
        target.mkdir()
        plan_copies([MatchResult(name="垫片", model=str(source))], target)
        assert os.listdir(target) == []


class TestExecuteCopies:
    def test_copies_file(self, tmp_path, make_files):
        (source,) = make_files(tmp_path / "src", ["垫片.STEP"])
        target = tmp_path / "dst"
        target.mkdir()
        report = execute_copies(plan_copies([MatchResult(name="垫片", model=str(source))], target))

        assert len(report.copied) == 1
        assert (target / "垫片.STEP").exists()
        assert report.failed == []

    def test_never_overwrites_same_name(self, tmp_path, make_files):
        (source,) = make_files(tmp_path / "src", ["垫片.STEP"], )
        target = tmp_path / "dst"
        (existing,) = make_files(target, ["垫片.STEP"])
        existing.write_text("人工改过的内容", encoding="utf-8")
        source.write_text("新的", encoding="utf-8")

        report = execute_copies(plan_copies([MatchResult(name="垫片", model=str(source))], target))

        assert report.skipped and not report.copied
        assert existing.read_text(encoding="utf-8") == "人工改过的内容"

    def test_records_failures_without_raising(self, tmp_path, make_files, monkeypatch):
        (source,) = make_files(tmp_path / "src", ["垫片.STEP"])
        action = CopyAction(
            name="垫片", src=str(source), dst=str(tmp_path / "dst" / "垫片.STEP"), kind="copy"
        )
        monkeypatch.setattr(
            "speedup.matching.shutil.copy2",
            lambda *a, **k: (_ for _ in ()).throw(PermissionError("被占用")),
        )
        report = execute_copies([action])
        assert report.failed and not report.copied
        assert "被占用" in report.failed[0][2]

    def test_counts_inplace_and_skipped(self, tmp_path, make_files):
        (source,) = make_files(tmp_path, ["垫片.STEP"])
        report = execute_copies(plan_copies([MatchResult(name="垫片", model=str(source))], tmp_path))
        assert len(report.inplace) == 1
        assert report.total_attempted == 1

    def test_empty_plan(self):
        report = execute_copies([])
        assert report.total_attempted == 0


class TestFindSharedMatches:
    def test_detects_two_names_sharing_one_file(self, scanned):
        (shared,) = scanned(["HEAD_FRONT.step"])
        file = ScannedFile(path=shared.path, stem="HEAD_FRONT", ext=".step")
        results = [
            MatchResult(name="HEAD_FRONT", model=file.path),
            MatchResult(name="HEAD_FRONT_456", model=file.path),
        ]
        clashes = find_shared_matches(results)
        assert len(clashes) == 1
        assert clashes[0][2] == "HEAD_FRONT.step"

    def test_no_clash_for_distinct_files(self, tmp_path, make_files):
        (one,) = make_files(tmp_path, ["a.step"])
        (two,) = make_files(tmp_path, ["b.step"])
        results = [
            MatchResult(name="a", model=str(one)),
            MatchResult(name="b", model=str(two)),
        ]
        assert find_shared_matches(results) == []


class TestUnusedFiles:
    def test_reports_unmatched_files(self, tmp_path, make_files):
        (used,) = make_files(tmp_path, ["垫片.step"])
        (unused,) = make_files(tmp_path, ["无关.step"])
        files = [
            ScannedFile(path=str(used), stem="垫片", ext=".step"),
            ScannedFile(path=str(unused), stem="无关", ext=".step"),
        ]
        leftovers = unused_files(files, [MatchResult(name="垫片", model=str(used))])
        assert [f.stem for f in leftovers] == ["无关"]

    def test_all_used(self, tmp_path, make_files):
        (only,) = make_files(tmp_path, ["垫片.step"])
        files = [ScannedFile(path=str(only), stem="垫片", ext=".step")]
        assert unused_files(files, [MatchResult(name="垫片", model=str(only))]) == []
