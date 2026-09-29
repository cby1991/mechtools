"""
测试组 B：纯逻辑行为测试。

这里喂的是**真实场景里出现过的字符串**（零件名、路径、PowerShell 回传值），
所以它同时起了"回归测试"和"示例文档"的作用 ——
想知道 U_HitKeyword 到底放行什么、挡住什么，看这里的用例最快。

全部离线可跑，不需要 SolidWorks。
"""

from __future__ import annotations

import pytest

import vba as V


# ==========================================================================
# 1. 路径处理
# ==========================================================================

class TestNormPath:
    def test_strips_trailing_backslash(self):
        assert V.norm_path("D:\\a\\b\\") == "D:\\a\\b"

    def test_strips_many_trailing_backslashes(self):
        assert V.norm_path("D:\\a\\\\\\") == "D:\\a"

    def test_keeps_drive_root(self):
        """盘根必须保留尾斜杠 —— 削成 'D:' 之后拼路径会变成 'D:a\\b'。"""
        assert V.norm_path("D:\\") == "D:\\"
        assert V.norm_path("C:\\") == "C:\\"

    def test_root_would_break_if_bound_changed(self):
        """反证：如果边界是 Len > 2，盘根就会被削坏。这里把后果写清楚。"""
        broken = "D:\\"
        while len(broken) > 2 and broken.endswith("\\"):
            broken = broken[:-1]
        assert broken == "D:"          # 这就是为什么必须是 Len > 3

    def test_trims_surrounding_spaces_only(self):
        """VBA 的 Trim$ 只去空格 —— Tab 不该被动。"""
        assert V.norm_path("  D:\\a  ") == "D:\\a"
        assert V.norm_path("\tD:\\a") == "\tD:\\a"

    def test_unc_path(self):
        assert V.norm_path("\\\\srv\\share\\dir\\") == "\\\\srv\\share\\dir"

    def test_unc_root_keeps_single_trailing_slash(self):
        """
        '\\\\srv\\' 长度 6 > 3 → 尾斜杠被削掉，磁盘上是『网络共享根』。
        注意：这条**不是**在说"UNC 根该保留斜杠" —— 而是如实记录 U_NormPath 的行为，
        顺带说明为什么 U_IsAbsPath 必须靠前两个字符判 UNC，而不是靠结尾。
        """
        assert V.norm_path("\\\\srv\\") == "\\\\srv"

    def test_unc_deep_path_strips(self):
        assert V.norm_path("\\\\srv\\share\\dir\\") == "\\\\srv\\share\\dir"

    def test_empty(self):
        assert V.norm_path("") == ""
        assert V.norm_path("   ") == ""


class TestJoinPath:
    def test_normal(self):
        assert V.join_path("D:\\out", "a.STEP") == "D:\\out\\a.STEP"

    def test_does_not_double_backslash(self):
        assert V.join_path("D:\\out\\", "a.STEP") == "D:\\out\\a.STEP"

    def test_empty_dir_returns_name_bare(self):
        """空目录时绝不能拼出 '\\a.STEP'（会变成"当前盘根"下的路径）。"""
        assert V.join_path("", "a.STEP") == "a.STEP"

    def test_drive_root(self):
        assert V.join_path("D:\\", "a.STEP") == "D:\\a.STEP"


class TestFileNameParentBaseExt:
    def test_file_name(self):
        assert V.file_name("D:\\a\\b\\c.SLDPRT") == "c.SLDPRT"
        assert V.file_name("c.SLDPRT") == "c.SLDPRT"
        assert V.file_name("") == ""

    def test_parent_normal(self):
        assert V.parent("D:\\a\\b\\c.SLDPRT") == "D:\\a\\b"

    def test_parent_at_drive_root_stops_at_three(self):
        """到 'D:\\a' 的父级应该是 'D:\\'，而不是 'D:'。"""
        assert V.parent("D:\\a") == "D:\\"

    def test_parent_no_backslash(self):
        assert V.parent("bare.SLDPRT") == ""

    def test_base_name(self):
        assert V.base_name("D:\\a\\b\\c.SLDPRT") == "c"
        assert V.base_name("D:\\a\\.hidden") == ".hidden"

    def test_base_name_multiple_dots(self):
        assert V.base_name("D:\\a\\hex.bolt.M8.SLDPRT") == "hex.bolt.M8"

    def test_ext_is_uppercased(self):
        assert V.ext("D:\\a\\c.sldprt") == ".SLDPRT"
        assert V.ext("D:\\a\\c.Step") == ".STEP"

    def test_ext_none(self):
        assert V.ext("D:\\a\\noext") == ""

    def test_dotfile_has_no_base_strip(self):
        """'.hidden' 的点在第 1 位，U_BaseName 要求 i > 1，不该被误当扩展名切掉。"""
        assert V.base_name(".hidden") == ".hidden"


class TestSafeName:
    @pytest.mark.parametrize("bad", list('\\/:*?"<>|'))
    def test_each_illegal_char(self, bad):
        assert V.safe_name(f"a{bad}b") == "a_b"

    def test_clean_name_untouched(self):
        assert V.safe_name("内六角螺钉M8x30") == "内六角螺钉M8x30"

    def test_real_world_folder_name(self):
        assert V.safe_name("D:\\图纸\\总成") == "D__图纸_总成"


class TestFileKey:
    def test_case_insensitive_dedup(self):
        a = V.file_key("D:\\a\\X.SLDPRT")
        b = V.file_key("d:\\A\\x.sldprt")
        assert a == b, "去重 key 必须大小写不敏感，否则同一文件会被导两遍"

    def test_trailing_slash_ignored(self):
        assert V.file_key("D:\\a\\X.SLDPRT\\") == V.file_key("D:\\a\\X.SLDPRT")


# ==========================================================================
# 2. U_IsAbsPath —— GetDocumentDependencies2 重复收录那个 bug 的防线
# ==========================================================================

class TestIsAbsPath:
    def test_drive_absolute(self):
        assert V.is_abs_path("D:\\a\\b.SLDPRT") is True

    def test_unc_absolute(self):
        assert V.is_abs_path("\\\\srv\\share\\b.SLDPRT") is True

    def test_relative_name_with_backslash_is_not_absolute(self):
        """
        【核心用例】GetDocumentDependencies2 返回的"名称"在有些保存方式下形如
        '.\\零件.SLDPRT' 或 '..\\零件.SLDPRT' —— 它们**含反斜杠但不是绝对路径**。
        如果只判"含不含 \\"，同一零件会被收两次 → 导出 零件.STEP 和 零件_2.STEP。
        """
        assert V.is_abs_path(".\\零件.SLDPRT") is False
        assert V.is_abs_path("..\\零件.SLDPRT") is False

    def test_bare_name_is_not_absolute(self):
        assert V.is_abs_path("零件.SLDPRT") is False

    def test_short_strings_rejected(self):
        assert V.is_abs_path("") is False
        assert V.is_abs_path("D") is False
        assert V.is_abs_path("D:") is False
        assert V.is_abs_path("\\\\") is False

    def test_colon_at_position_2_only(self):
        """'D:\\x' 算绝对；'xD:\\x' 不算（冒号在第 3 位）。"""
        assert V.is_abs_path("xD:\\x") is False


# ==========================================================================
# 3. U_HitKeyword —— 标准件过滤的全部真实用例
# ==========================================================================

class TestHitKeywordChineseSubstring:
    """中文关键字走子串匹配：必须能命中"紧贴尺寸代号"的写法。"""

    @pytest.mark.parametrize(
        "hay",
        ["内六角螺钉M8x30", "内六角螺钉", "M8内六角螺钉", "内六角螺钉_M8x30"],
    )
    def test_chinese_keyword_hits_anywhere(self, hay):
        assert V.hit_keyword(hay, "螺钉") is True

    def test_chinese_does_not_hit_unrelated(self):
        assert V.hit_keyword("支架", "螺钉") is False

    def test_mixed_keyword_still_substring(self):
        """『内六角螺钉』整体作关键字时，也要能命中后面跟尺寸的写法。"""
        assert V.hit_keyword("内六角螺钉M8x30", "内六角螺钉") is True


class TestHitKeywordSeparatorSubstring:
    """含 / - _ 空格 的关键字自带边界，直接子串。"""

    @pytest.mark.parametrize(
        "needle", ["GB/T", "GB-T", "Spring Pin", "O-Ring", "Snap Ring", "DIN_912"]
    )
    def test_separator_keywords_substring(self, needle):
        assert V.hit_keyword(f"x{needle}y", needle) is True

    def test_gb_slash_real_world(self):
        assert V.hit_keyword("GB_T 5783-2016 螺栓", "GB/T") is False  # 下划线≠斜杠
        assert V.hit_keyword("GB/T5783-2016 螺栓", "GB/T") is True


class TestHitKeywordWordBoundaryShouldMiss:
    """以下都必须**挡住** —— 这些是"词中词"误伤，每一个都对应过一次真实的误过滤。"""

    @pytest.mark.parametrize(
        "hay,needle",
        [
            ("Spindle", "Pin"),        # Pin 藏在 Spindle 里
            ("Walnut", "Nut"),         # Nut 藏在 Walnut 里
            ("ISOLATION", "ISO"),      # ISO 是 ISOLATION 的前缀
            ("Bolted", "Bolt"),        # 后缀是字母 e
            ("Screwdriver", "Screw"),  # 后缀是字母 d
            ("Nutshell", "Nut"),       # 后缀是字母 s
            ("aNut", "Nut"),           # Nut 前面是字母
            ("Nutcase", "Nut"),        # 后缀是字母 c
        ],
    )
    def test_word_in_word_is_blocked(self, hay, needle):
        assert V.hit_keyword(hay, needle) is False, f"{hay} 不该命中 {needle}"

    def test_case_insensitive_but_still_blocked(self):
        assert V.hit_keyword("sPiNdLe", "PIN") is False


class TestHitKeywordWordBoundaryShouldHit:
    """以下都必须**放行** —— 关键字紧贴数字是最常见的合法写法，不能漏。"""

    @pytest.mark.parametrize(
        "hay,needle",
        [
            ("GBT5782", "GBT"),        # 后面紧跟数字 → 放行
            ("ISO4762", "ISO"),        # 后面紧跟数字 → 放行
            ("Nut8", "Nut"),           # 后面紧跟数字 → 放行
            ("Pin_M8", "Pin"),         # 后面是下划线（非字母）→ 放行
            ("x Nut", "Nut"),          # 前面是空格
            ("(Nut)", "Nut"),          # 前面是括号
            ("-Nut", "Nut"),           # 前面是连字符
            ("Nut", "Nut"),            # 整串就是关键字
        ],
    )
    def test_valid_writings_are_found(self, hay, needle):
        assert V.hit_keyword(hay, needle) is True, f"{hay} 应该命中 {needle}"


class TestHitKeywordMultiOccurrence:
    """一次字符串里出现多次关键字：第一次被挡、第二次合格，也得能命中。"""

    def test_second_occurrence_wins(self):
        #          0123456
        hay = "Spindle Pin"       # 第一个 Pin 在 Spindle 里（挡），第二个独立（放行）
        assert V.hit_keyword(hay, "Pin") is True

    def test_all_occurrences_blocked(self):
        assert V.hit_keyword("Spindle Spinner", "Pin") is False

    def test_occurrence_at_end_of_string(self):
        """关键字紧贴字符串末尾：右边界为空串 → 必须放行。"""
        assert V.hit_keyword("Hex Nut", "Nut") is True

    def test_occurrence_at_start_of_string(self):
        """关键字在开头：左边界为空串 → 必须放行。"""
        assert V.hit_keyword("Nut8", "Nut") is True


class TestHitKeywordEdges:
    def test_empty_inputs(self):
        assert V.hit_keyword("", "Pin") is False
        assert V.hit_keyword("Pin", "") is False
        assert V.hit_keyword("", "") is False

    def test_keyword_longer_than_haystack(self):
        assert V.hit_keyword("Nut", "Nutter") is False


# ==========================================================================
# 4. U_KeywordList —— CFG_* 逗号表解析
# ==========================================================================

class TestKeywordList:
    def test_splits_and_trims(self):
        assert V.keyword_list("a, b ,c") == ["a", "b", "c"]

    def test_drops_empty_entries(self):
        assert V.keyword_list("a,,b, ,c") == ["a", "b", "c"]

    def test_only_commas(self):
        assert V.keyword_list(",,,,") == []

    def test_empty(self):
        assert V.keyword_list("") == []

    def test_all_whitespace(self):
        assert V.keyword_list("   ") == []

    def test_preserves_chinese(self):
        assert V.keyword_list("螺钉,螺母,垫圈") == ["螺钉", "螺母", "垫圈"]


# ==========================================================================
# 5. U_CleanResult —— PowerShell 回传值清洗
# ==========================================================================

class TestCleanResult:
    def test_strips_crlf(self):
        """选文件夹回传的路径末尾带 CRLF —— 不清掉拿去建目录必然失败。"""
        assert V.clean_result("D:\\out\\\r\n") == "D:\\out\\"

    def test_strips_lone_cr_and_lf(self):
        assert V.clean_result("D:\\out\\\r") == "D:\\out\\"
        assert V.clean_result("D:\\out\\\n") == "D:\\out\\"

    def test_strips_tab(self):
        assert V.clean_result("D:\\out\\\t") == "D:\\out\\"

    def test_strips_leading_bom(self):
        assert V.clean_result("\ufeffD:\\out") == "D:\\out"

    def test_cancel_preserved_as_single_char(self):
        """【关键】'C'（取消）清洗后必须仍是 'C'，不能被当成空结果。"""
        assert V.clean_result("C\r\n") == "C"
        assert V.clean_result("C") == "C"

    def test_mask_preserved(self):
        """三位掩码不能被 Trim 吃掉。"""
        assert V.clean_result("010\r\n") == "010"
        assert V.clean_result("111\n") == "111"

    def test_strips_leading_bom_from_mask(self):
        """带 BOM 的掩码 —— 不去 BOM 的话 Mid$(s,1,1) 会拿到 BOM 而不是 '0'/'1'。"""
        assert V.clean_result("\ufeff010") == "010"

    def test_empty(self):
        assert V.clean_result("") == ""
        assert V.clean_result("\r\n") == ""

    def test_only_spaces_trimmed_not_inner(self):
        assert V.clean_result("  D:\\a b  ") == "D:\\a b"


class TestAskSkipMaskDecoding:
    """
    U_AskSkipOptions 对清洗后字符串的三级判定 —— 这里把它的语义完整复现一遍。
    返回：-1 = 取消；-2 = 对话框不可用（按默认走）；0..7 = 位掩码。
    """

    @staticmethod
    def decode(raw: str) -> int:
        s = V.clean_result(raw)
        if s == "C":
            return -1
        if len(s) == 3:
            n = 0
            if V.mid(s, 1, 1) == "1":
                n += 1
            if V.mid(s, 2, 1) == "1":
                n += 2
            if V.mid(s, 3, 1) == "1":
                n += 4
            return n
        return -2

    @pytest.mark.parametrize(
        "raw,expected",
        [
            ("000", 0), ("100", 1), ("010", 2), ("001", 4),
            ("110", 3), ("101", 5), ("011", 6), ("111", 7),
            ("000\r\n", 0), ("\ufeff111", 7),
        ],
    )
    def test_valid_masks(self, raw, expected):
        assert self.decode(raw) == expected

    def test_cancel(self):
        assert self.decode("C") == -1
        assert self.decode("C\r\n") == -1

    def test_cancel_is_not_swallowed_by_length_check(self):
        """
        反证：如果先做"长度必须是 3"的校验，取消（长度 1）会落到 -2，
        调用方按默认继续 → 用户按了取消却还在导出。v2.2.0~v2.2.2 的原 bug。
        """
        s = V.clean_result("C\r\n")
        assert len(s) != 3
        # 正确顺序下，先命中取消分支，所以这里必须是 -1
        assert self.decode("C\r\n") == -1

    @pytest.mark.parametrize("raw", ["", "E", "0", "00", "0000", "   "])
    def test_bad_results_fall_back_to_default(self, raw):
        """长度不为 3、且不是 'C' → -2（按默认继续），绝不能让一个界面问题卡死整个导出。"""
        assert self.decode(raw) == -2

    def test_length3_garbage_is_accepted_as_mask(self):
        """
        特征测试（记录现状，不是赞美）：源码只判 `Len(sResp) = 3`，**不校验内容**，
        所以 'X1Y' 这种三字符垃圾会被当成掩码 010 → 返回 2。
        影响可控：正常路径下脚本只写 0/1 或 'C'，走到这里说明脚本被改坏了，
        后果是"多导出一类"而不是"丢数据"。这里如实钉住，免得将来误以为有校验。
        """
        assert self.decode("X1Y") == 2
        assert self.decode("abc") == 0      # 三个字符都不是 '1' → 掩码 0

    def test_length3_with_leading_one_still_parsed(self):
        """'1ab' → bit0 命中 → 1。同样是"只按位置取字符"的后果。"""
        assert self.decode("1ab") == 1


# ==========================================================================
# 6. U_PsQuote / U_Csv 转义
# ==========================================================================

class TestPsQuote:
    def test_plain(self):
        assert V.ps_quote("D:\\out") == "'D:\\out'"

    def test_internal_single_quote_doubled(self):
        """PowerShell 单引号字面量里，单引号要用两个表示。"""
        assert V.ps_quote("it's") == "'it''s'"

    def test_all_quotes(self):
        assert V.ps_quote("'''") == "''''''''"

    def test_backslash_and_dollar_untouched(self):
        """单引号字面量里 $ 和 \\ 不转义 —— 这正是选它而不是双引号的原因。"""
        assert V.ps_quote("$env:TEMP\\a") == "'$env:TEMP\\a'"

    def test_chinese_path(self):
        assert V.ps_quote("D:\\图纸\\总成") == "'D:\\图纸\\总成'"

    def test_empty(self):
        assert V.ps_quote("") == "''"


class TestCsvCell:
    def test_plain(self):
        assert V.csv_cell("abc") == '"abc"'

    def test_internal_double_quote_doubled(self):
        assert V.csv_cell('a"b') == '"a""b"'

    def test_chinese(self):
        assert V.csv_cell("已跳过") == '"已跳过"'

    def test_empty(self):
        assert V.csv_cell("") == '""'


# ==========================================================================
# 7. 跳过掩码的暂存 → 合并
# ==========================================================================

class TestKindSkipped:
    @pytest.mark.parametrize(
        "kind,env,sup,hid,expected",
        [
            (1, True, False, False, True),
            (1, False, True, True, False),
            (2, False, True, False, True),
            (2, True, False, True, False),
            (4, False, False, True, True),
            (4, True, True, False, False),
        ],
    )
    def test_category_mapped_to_right_switch(self, kind, env, sup, hid, expected):
        assert V.kind_skipped(kind, env, sup, hid) == expected

    def test_unknown_kind_is_false(self):
        """0（正常）或任何意外值都必须返回 False —— 不能被当成"要跳过"。"""
        assert V.kind_skipped(0, True, True, True) is False
        assert V.kind_skipped(3, True, True, True) is False
        assert V.kind_skipped(99, True, True, True) is False


class TestMergeOne:
    def test_checked_category_restores_model(self):
        """用户勾了"也导出已隐藏" → 该零件要回到待导出列表。"""
        models, results = [], []
        V.merge_one([("Part1", "D:\\p1.SLDPRT", "已隐藏")], 4, 4, "已隐藏", models, results)
        assert models == ["D:\\p1.SLDPRT"]
        assert results == [], "放回来后不该再写跳过记录"

    def test_unchecked_category_records_skip(self):
        models, results = [], []
        V.merge_one([("Part1", "D:\\p1.SLDPRT", "已隐藏")], 4, 0, "已隐藏", models, results)
        assert models == []
        assert results == [("SKIP", "Part1", "已隐藏")]

    def test_checked_but_pathless_is_reported_honestly(self):
        """压缩组件常常拿不到磁盘路径 —— 勾了也导不出来，必须如实写明。"""
        models, results = [], []
        V.merge_one([("Asm1", "", "已压缩")], 2, 2, "已压缩", models, results)
        assert models == []
        assert len(results) == 1
        assert "取不到文件路径" in results[0][4]

    def test_no_contradictory_record_for_shared_model(self):
        """
        【核心用例】同一零件两个实例：一个隐藏（暂存）、一个可见（已进待导出）。
        用户不勾"导出已隐藏" → 不能写"跳过"，否则清单里一边成功一边跳过，自相矛盾。
        """
        models = ["D:\\p1.SLDPRT"]
        results = []
        V.merge_one([("Part1", "D:\\p1.SLDPRT", "已隐藏")], 4, 0, "已隐藏", models, results)
        assert results == [], "已被别的实例导出过的零件，不该再记一条跳过"
        assert models == ["D:\\p1.SLDPRT"]

    def test_shared_model_match_is_case_insensitive(self):
        """Windows 路径不分大小写，匹配也必须不分。"""
        models = ["D:\\P1.SLDPRT"]
        results = []
        V.merge_one([("Part1", "d:\\p1.sldprt", "已隐藏")], 4, 0, "已隐藏", models, results)
        assert results == []

    def test_unrelated_pathless_skip_still_recorded(self):
        """路径取不到、又没人勾它 —— 照样要记跳过（这是最常见的正常情况）。"""
        models, results = [], []
        V.merge_one([("Asm1", "", "已压缩")], 2, 0, "已压缩", models, results)
        assert results == [("SKIP", "Asm1", "已压缩")]

    def test_mixed_batch(self):
        models = ["D:\\shared.SLDPRT"]
        results = []
        rows = [
            ("Shared", "D:\\shared.SLDPRT", "已隐藏"),   # 已被导出 → 不写
            ("Lonely", "D:\\lonely.SLDPRT", "已隐藏"),   # 没进列表、没勾 → 记跳过
            ("NoPath", "", "已隐藏"),                    # 没路径、没勾 → 记跳过
        ]
        V.merge_one(rows, 4, 0, "已隐藏", models, results)
        assert results == [
            ("SKIP", "Lonely", "已隐藏"),
            ("SKIP", "NoPath", "已隐藏"),
        ]

    def test_duplicate_restore_does_not_add_twice(self):
        """勾了两类、同一个路径出现两次 → 待导出列表里仍只能有一份。"""
        models, results = [], []
        p = [("P", "D:\\p.SLDPRT", "已隐藏")]
        V.merge_one(p, 4, 4, "已隐藏", models, results)
        V.merge_one(p, 4, 4, "已隐藏", models, results)
        assert models == ["D:\\p.SLDPRT"]


class TestFullMaskFlow:
    """把三类一起过一遍，模拟真实的"扫完 → 询问 → 合并"全流程。"""

    def test_bitmask_124_selects_independently(self):
        env = [("E1", "D:\\e1.SLDPRT", "封套")]
        sup = [("S1", "D:\\s1.SLDPRT", "已压缩")]
        hid = [("H1", "D:\\h1.SLDPRT", "已隐藏")]

        # 只勾"封套 + 已隐藏" → bit0(封套=1) + bit2(已隐藏=4) = 5
        mask = 1 | 4
        models, results = [], []
        V.merge_one(env, 1, mask, "封套", models, results)
        V.merge_one(sup, 2, mask, "已压缩", models, results)
        V.merge_one(hid, 4, mask, "已隐藏", models, results)

        assert sorted(models) == ["D:\\e1.SLDPRT", "D:\\h1.SLDPRT"]
        assert results == [("SKIP", "S1", "已压缩")]


# ==========================================================================
# 8. G_SkipKind 的父级链遍历
# ==========================================================================

class TestSkipKindChain:
    def test_all_normal(self):
        kind, why, _ = V.skip_kind_chain([0, 0, 0])
        assert kind == 0 and why == ""

    def test_self_bad_no_parent_suffix(self):
        kind, why, hops = V.skip_kind_chain([4])
        assert kind == 4
        assert why == "已隐藏", "自身就不正常时不该加「上级装配体」后缀"
        assert hops == 1

    def test_parent_bad_gets_suffix(self):
        """自身正常、子装配体（父级）被隐藏 —— 实际同样看不见，要跳过并标明是上级导致。"""
        kind, why, hops = V.skip_kind_chain([0, 4])
        assert kind == 4
        assert why == "已隐藏（上级装配体）"
        assert hops == 2

    def test_first_bad_in_chain_wins(self):
        """逐级往上，第一个不正常的就是结论。"""
        kind, why, _ = V.skip_kind_chain([0, 0, 2, 4])
        assert kind == 2
        assert why == "已压缩（上级装配体）"

    def test_envelope_checked_before_hidden_on_same_comp(self):
        """单个组件同时满足多个类别时，按 封套 → 压缩 → 隐藏 归类（这里模拟已归类的结果）。"""
        kind, _, _ = V.skip_kind_chain([1])
        assert kind == 1

    def test_cycle_guard_bails_out(self):
        """父级链异常长（成环）时，必须在 64 次后放弃，不能挂死 SolidWorks。"""
        kind, why, hops = V.skip_kind_chain([0] * 100, limit=64)
        assert kind == 0, "防环 bail 时按「不跳过」处理"
        assert hops == 64, f"应该在 64 跳后停下，实际 {hops}"

    def test_bad_state_found_within_limit(self):
        states = [0] * 10 + [4] + [0] * 100
        kind, why, hops = V.skip_kind_chain(states, limit=64)
        assert kind == 4
        assert hops == 11


# ==========================================================================
# 9. 单字符判定（AscW 边界的回归）
# ==========================================================================

class TestAsciiClassifiers:
    @pytest.mark.parametrize("c", list("0123456789"))
    def test_digits(self, c):
        assert V.is_ascii_alnum(c) is True
        assert V.is_ascii_alpha(c) is False

    @pytest.mark.parametrize("c", list("abcXYZ"))
    def test_letters(self, c):
        assert V.is_ascii_alnum(c) is True
        assert V.is_ascii_alpha(c) is True

    @pytest.mark.parametrize("c", ["_", "-", ".", " ", "/", "\\", "(", ")"])
    def test_punctuation_is_neither(self, c):
        assert V.is_ascii_alnum(c) is False
        assert V.is_ascii_alpha(c) is False

    @pytest.mark.parametrize("c", ["螺", "钉", "中", "é", "①"])
    def test_non_ascii_is_neither(self, c):
        assert V.is_ascii_alnum(c) is False
        assert V.is_ascii_alpha(c) is False

    def test_empty_string(self):
        assert V.is_ascii_alnum("") is False
        assert V.is_ascii_alpha("") is False

    def test_del_character_is_not_alnum(self):
        """0x7F (DEL) 紧跟在 'z' 之后，最容易因为边界写成 <= 而误判。"""
        assert V.is_ascii_alnum("\x7f") is False
        assert V.is_ascii_alpha("\x7f") is False

    def test_boundary_chars(self):
        """每个区间的上下沿字符：'/' 之后是 '0'，':' 之前是 '9'，等等。"""
        for c in ["/", ":", "@", "[", "`", "{", "}", "~"]:
            assert V.is_ascii_alnum(c) is False, f"{c!r} (U+{ord(c):04X}) 不该算字母数字"
            assert V.is_ascii_alpha(c) is False, f"{c!r} (U+{ord(c):04X}) 不该算字母"

    def test_range_edges_are_inclusive(self):
        """边界值本身必须在区间内：'0' '9' 'A' 'Z' 'a' 'z'。"""
        for c in ["0", "9", "A", "Z", "a", "z"]:
            assert V.is_ascii_alnum(c) is True, f"{c!r} 必须在区间内"
        for c in ["A", "Z", "a", "z"]:
            assert V.is_ascii_alpha(c) is True, f"{c!r} 必须在区间内"


class TestHasNonAscii:
    def test_pure_ascii(self):
        assert V.has_non_ascii("ABC-123_x") is False

    def test_del_is_treated_as_non_ascii_by_source(self):
        """
        特征测试：源码判的是 `u > 126`，而 0x7F (DEL) = 127 > 126，
        所以 DEL **会**被算成"非 ASCII"。这只是个理论边角（Windows 文件名里不会有 DEL），
        但既然源码如此，测试就如实记录 —— 免得将来有人"顺手改成 >= 127"却不知道动了什么。
        """
        assert V.has_non_ascii("\x7f") is True

    def test_max_true_ascii(self):
        """0x7E ('~') 是源码规则下最后一个"真 ASCII"。"""
        assert V.has_non_ascii("~") is False
        assert V.has_non_ascii("\x7e") is False

    def test_chinese(self):
        assert V.has_non_ascii("螺钉") is True

    def test_mixed(self):
        assert V.has_non_ascii("M8螺钉") is True

    def test_empty(self):
        assert V.has_non_ascii("") is False

    def test_latin1_accent(self):
        assert V.has_non_ascii("é") is True

    def test_cjk_extension_needs_negative_wrap(self):
        """U+8000 以上的字符：VBA 的 AscW 返回负数，靠 +65536 掰回来才 > 126。"""
        for ch in ["\u8000", "\u9fa5", "\uffe5"]:
            assert V.ascw(ch) < 0, f"{ch!r} 的 AscW 应该是负数（模拟 VBA 行为）"
            assert V.has_non_ascii(ch) is True, f"{ch!r} 必须被认出是非 ASCII"


class TestHasSeparator:
    @pytest.mark.parametrize("s", ["GB/T", "GB-T", "DIN_912", "Spring Pin", "a.b", "O-Ring"])
    def test_with_separators(self, s):
        assert V.has_separator(s) is True

    @pytest.mark.parametrize("s", ["GBT", "ISO4762", "Nut", "M8"])
    def test_without_separators(self, s):
        assert V.has_separator(s) is False

    def test_chinese_counts_as_separator(self):
        """中文不是 ASCII 字母数字 → 含中文的关键字一律走子串匹配。"""
        assert V.has_separator("螺钉") is True


# ==========================================================================
# 10. 模型 / 工程图识别
# ==========================================================================

class TestDocClassify:
    @pytest.mark.parametrize("name", ["a.SLDPRT", "a.sldprt", "a.SldPrt", "a.SLDASM"])
    def test_models(self, name):
        assert V.is_model(f"D:\\x\\{name}") is True
        assert V.is_drawing(f"D:\\x\\{name}") is False

    @pytest.mark.parametrize("name", ["a.SLDDRW", "a.slddrw"])
    def test_drawings(self, name):
        assert V.is_drawing(f"D:\\x\\{name}") is True
        assert V.is_model(f"D:\\x\\{name}") is False

    def test_step_is_neither(self):
        assert V.is_model("D:\\x\\a.STEP") is False
        assert V.is_drawing("D:\\x\\a.STEP") is False

    def test_no_extension(self):
        assert V.is_model("D:\\x\\a") is False


# ==========================================================================
# 11. 输出重名序号（U_TargetPath 的本次运行内去重）
# ==========================================================================

class TestTargetPathNaming:
    @staticmethod
    def target(base: str, ext_: str, used: set) -> str:
        """
        复现 U_TargetPath 的"同一次运行内重名加序号"逻辑（不含"文件已存在则跳过"分支）。
        """
        try_ = V.join_path("D:\\out", base + ext_)
        if V.ucase(try_) in used:
            for i in range(2, 1000):
                try_ = V.join_path("D:\\out", f"{base}_{i}{ext_}")
                if V.ucase(try_) not in used:
                    break
        used.add(V.ucase(try_))
        return try_

    def test_first_gets_plain_name(self):
        used = set()
        assert self.target("part", ".STEP", used) == "D:\\out\\part.STEP"

    def test_second_gets_suffix_2(self):
        used = set()
        self.target("part", ".STEP", used)
        assert self.target("part", ".STEP", used) == "D:\\out\\part_2.STEP"

    def test_third_gets_suffix_3(self):
        used = set()
        for _ in range(3):
            last = self.target("part", ".STEP", used)
        assert last == "D:\\out\\part_3.STEP"

    def test_case_insensitive_collision(self):
        """'Part.STEP' 和 'part.step' 在 Windows 上是同一个文件，必须也加序号。"""
        used = set()
        self.target("Part", ".STEP", used)
        assert self.target("part", ".STEP", used) == "D:\\out\\part_2.STEP"

    def test_不同扩展名不冲突(self):
        used = set()
        self.target("part", ".STEP", used)
        assert self.target("part", ".PDF", used) == "D:\\out\\part.PDF"
