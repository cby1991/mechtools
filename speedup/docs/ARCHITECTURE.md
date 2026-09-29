# 架构说明

本文回答两个问题：**代码怎么分层**，以及**为什么这么分**。
行为准则见 [`AGENTS.md`](../AGENTS.md)，使用说明见 [`README.md`](../README.md)。

---

## 1. 分层总览

```
  双击 .lnk          命令行
  （pythonw.exe）    （uv run speedup）
        │                 │
        └────────┬────────┘
                 ▼
   ┌─────────────────────────────┐   ┌──────────────────────────┐
   │  cli.py                     │   │  shortcut.py             │
   │  ├─ main()  入口与分发       │   │  生成 .lnk（仅安装时用一次）│
   │  ├─ run_gui()  开图形界面    │   └──────────────────────────┘
   │  ├─ run_menu() 文字菜单      │
   │  ├─ run_doctor() / archive() │
   │  └─ run_shortcut()           │
   └──────┬───────────────┬──────┘
          │               │
          │  不开窗口时    │  开窗口时
          ▼               ▼
   ┌────────────┐  ┌──────────────────────────────┐
   │ （直接调）  │  │  gui.py                      │
   │            │  │  ├─ SpeedupApp  主窗口/导航   │
   │            │  │  ├─ ToolPage    一页一工具    │
   │            │  │  │   只做「收参数 → 拼 argv」 │
   │            │  │  └─ 日志走 queue 回主线程     │
   └─────┬──────┘  └──────────────┬───────────────┘
         │                        │
         └────────────┬───────────┘
                      │ 三条入口汇合到同一条路：
                      │ build_parser() → TOOLS[name].run(args)
         ┌────────────▼─────────────────┐
         │  tools/   一个工具一个模块      │
         │  ├─ gbbuild.py  生成补料清单   │
         │  └─ gbcopy.py   匹配拷贝      │
         └──────────────┬───────────────┘
                        │ 只编排，不实现细节
        ┌───────────────┼──────────────────────────┐
        │               │                          │
┌───────▼────────┐  ┌───▼────────────────┐  ┌──────▼───────────┐
│ 领域逻辑（纯）  │  │ IO 适配层           │  │ 基础设施          │
│ material.py    │  │ excel/reader.py    │  │ console.py       │
│ matching.py    │  │ excel/writer.py    │  │ dialogs.py       │
│ models.py      │  │ osutil.py          │  │ errors.py        │
└───────┬────────┘  └──────────┬─────────┘  │ config.py        │
        │                      │            └──────────────────┘
        └──────────┬───────────┘
                   ▼
          tests/  （直接测纯逻辑，无需造真实文件）
```

**关键**：`gui.py` 不在业务链上，它和 `cli.py` 是**同级的入口**。
界面里没有一行业务逻辑，只有「把控件值拼成 argv」。这条约定见
[`AGENTS.md` R19](../AGENTS.md#r19--界面不许有自己的一套逻辑只许收参数--拼-argv--交给同一个-run-2026-09-28)。

---

## 2. 三条核心设计原则

### 2.1 纯函数优先：决策与执行分离

这是本项目**最重要**的结构约定，也是可测性的来源。

以 `matching.py` 为例：

| 函数 | 碰磁盘吗 | 干什么 |
| --- | --- | --- |
| `normalize` / `match_score` | ❌ | 打分 |
| `find_best` / `build_matches` | ❌ | 选出最佳匹配、定等级 |
| `plan_copies` | ❌ | **决定**每个文件该拷/该跳/本来就在 |
| `scan_files` | 读 | 扫描目录 |
| `execute_copies` | **写** | 真正执行拷贝 |

好处立刻体现：

- `--dry-run` 的实现 = 「只调 `plan_copies`，不调 `execute_copies`」，一行
- 测试可以精确断言「目标已有同名文件时必须判为 `skip_same_name`」，
  不用真的造两个目录再观察副作用
- 「不覆盖同名文件」这条安全约定有了**单元测试**，不再靠人肉 review

`material.py` 同理：

| 函数 | 依赖 |
| --- | --- |
| `detect_material` / `find_material_in_words` | 无依赖，纯字符串/列表运算 |
| `find_material_in_pdf` | 需要 pdfplumber |

于是**材料识别规则的测试不需要装 pdfplumber，也不需要真 PDF**。
只有「读 PDF」那一层是薄壳。

### 2.2 常量集中：`config.py` 是唯一真相

旧脚本的 `MODEL_EXTS`、`NAME_COL`、`HEADER_ROW` 散在各文件头部。
补料清单的列号在**生成侧**（writer）和**读取侧**（reader）各写一份 —— 改列就出事。

现在全部在 `config.py`，业务模块里不允许出现裸数字：

```python
# ❌ 旧写法
ws.cell(row=r, column=3, value=name)

# ✅ 现在
from .config import COL_NAME
ws.cell(row=r, column=COL_NAME, value=name)
```

**测试也从 `config` 取常量**，避免出现「实现改了列号、测试还断言旧值、
于是永远绿」的假安全感。

### 2.3 错误分层：可预期的失败不吐 traceback

```
SpeedupError（基类）
├── UserCancelled          用户关掉了窗口 / 输入空行
├── InputError             输入不符合预期（没有模型文件、找不到「名称」列）
├── UnsupportedFormatError 传了 .xls 老格式
├── MissingDependencyError 缺第三方库
└── ConfigError            配置文件非法
```

`cli.main()` 统一捕获前三类 → 打印 `[失败] 中文说明` → 返回退出码 2。
**未预期的异常让它冒泡**，这样开发期能立刻看到 traceback，不会被 `except: pass` 吞掉。

`osutil.py` 里的「顺手打开文件」是特例：失败只返回 `False`，
因为它是附加功能，不该让主流程挂掉。

---

## 3. 一次调用的完整链路

以 `uv run speedup gbcopy "清单.xlsx" "D:\库" --dry-run` 为例：

```
1. cli.main(["gbcopy", "清单.xlsx", "D:\库", "--dry-run"])
   └─ console.ensure_safe_output()          # 先保证中文输出安全
   └─ build_parser() 解析出 args
   └─ module = TOOLS["gbcopy"]              # 按名字取模块，不提前绑死函数
   └─ module.run(args)

2. gbcopy.run(args)
   ├─ excel.reader.extract_names(清单)      # 读「名称」列（只认一张工作表，见 AGENTS.md R24）
   │    └─ 选表：--sheet 指定 > 名为「补料清单」的表 > 第一张候选
   │    └─ read_workbook → openpyxl，失败则 read_xlsx_builtin（零依赖回退）
   ├─ matching.scan_files(搜索目录)          # 扫 .stp/.step/.pdf
   ├─ matching.build_matches(names, files)  # 纯函数 → MatchResult 列表
   │    └─ 每个结果带 level: ok / fuzzy / partial / miss
   ├─ gbcopy.render_match_table(...)        # console.render_table 排版
   ├─ matching.plan_copies(results, 目标)    # 纯函数 → CopyAction 列表
   └─ if args.dry_run:
        gbcopy._preview(actions)            # 只分类，不落盘
      else:
        matching.execute_copies(actions)    # 真正 shutil.copy2
   └─ gbcopy.render_report(...)             # 总结文本
```

**注意第 1 步**：`TOOLS["gbcopy"]` 而不是 `args._run`。
如果让 argparse 用 `set_defaults(_run=run)` 提前绑死函数对象，
菜单、命令行、测试就会走三条不同路径 —— 这个坑见
[`AGENTS.md` P3](../AGENTS.md#p3--parserset_defaults_runrun-把函数对象提前绑死)。

### 3.1 从图形界面出发的同一条路

在界面里点「开始执行」时，走的是：

```
1. ToolPage.on_run()
   └─ build_argv()                    # 纯取值，不碰窗口以外的东西
        → ["D:\某清单.xlsx", "D:\库", "--min-score", "3"]   # 位置参数在前
   └─ app.run_task(page, job)         # 交给后台线程

2. SpeedupApp._worker()   ← 子线程
   └─ redirect_stdout/_stderr 到 _QueueWriter
   └─ job() → ToolPage._invoke(argv)

3. ToolPage._invoke(argv)
   └─ cli.build_parser().parse_args(["gbcopy", *argv])
   └─ TOOLS[name].run(args)           # ←── 和命令行第 1 步完全同一行代码

4. 主线程每 80ms 一次 _drain()
   └─ 队列里的日志批量插进 Text 控件
   └─ 收到 ("done", code, kind, 秒数) 就解锁界面、刷状态栏
```

两个必须记住的点：

* **`print` 是输出通道**。工具的 `console.heading()` / `print_table()` 照常用，
  界面负责重定向和搬运 —— 所以工具完全不需要知道自己在不在界面里。
* **工作线程里不许碰控件**（Tk 不是线程安全的）。所以路径都在界面上先选好，
  工具内部那条「弹窗 → 手输」回退在图形界面下不会被走到。

---

## 4. 模块职责速查

| 模块 | 职责 | 依赖外部库 |
| --- | --- | --- |
| `cli.py` | 参数解析、子命令分发、图形界面/菜单/快捷方式入口、统一错误出口 | argparse |
| `gui.py` | tkinter 图形界面（分页、表单、日志搬运、后台任务调度） | tkinter（stdlib） |
| `shortcut.py` | 生成 `.lnk`（安装时用一次） | ctypes（stdlib） |
| `config.py` | **所有**常量 | 无 |
| `console.py` | 中文宽度对齐表格、分隔线、状态提示、`ensure_streams()` | 无 |
| `dialogs.py` | tkinter 弹窗 + 控制台手输回退 | tkinter（stdlib） |
| `errors.py` | 异常体系 | 无 |
| `material.py` | 材料识别规则 + PDF 读取 | pdfplumber（仅 `find_material_in_pdf`） |
| `matching.py` | 匹配打分 + 拷贝规划与执行 | 无（stdlib 的 `os`/`shutil`） |
| `models.py` | 领域数据结构（`SupplementEntry`） | 无 |
| `osutil.py` | 打开文件 / 定位文件夹 / 桌面目录 | 无（Windows 上读一次注册表） |
| `excel/reader.py` | 读 xlsx、抽「名称」列 | openpyxl（可回退到纯 stdlib） |
| `excel/writer.py` | 生成补料清单 | openpyxl |
| `tools/*.py` | 编排，不含算法 | — |

**依赖方向永远是单向的**：`gui / cli → tools → 领域/IO/基础设施`。
领域模块之间不互相依赖，也不反向依赖 `cli`。

> `shortcut.py` 是**安装期**的旁支，不在运行时链路上：它只在
> `uv run speedup shortcut` 时跑一次，产出物就是那个 `.lnk` 文件。

---

## 5. 关键取舍记录

设计决策连同**为什么不选另一条路**一起记在这里，防止后人「优化」掉。

### 5.1 用 tkinter，而不是继续用 PowerShell

| 方案 | 优点 | 缺点 | 结论 |
| --- | --- | --- | --- |
| PowerShell + WinForms | 窗口是系统原生样式 | 进程冷启动卡顿、临时 `.ps1` 被杀软误报、引号易被剥、无法单测 | ❌ |
| **tkinter** | stdlib、零依赖、可单测、跨平台 | 目录选择框不是资源管理器同款大窗口 | ✅ 采用 |
| ctypes + Win32 `IFileOpenDialog` | 真·资源管理器同款大窗口 | COM 代码复杂、难自动化测试 | ❌ 收益不抵复杂度 |

**遗留代价**：弹窗外观由 Tk 决定。已在 README 与 AGENTS.md 明确记录。

### 5.2 锁 Python 3.12

工具要 `tkinter`，而 uv 自带的 Python 3.12 分发包包含 tkinter（Tk 8.6），
3.13 的独立分发包不一定带。锁 3.12 最省心。锁定方式是 `.python-version`。

### 5.3 不强制 stdout 为 UTF-8

中文 Windows 控制台是 cp936，改 Python 侧编码 → 编码不匹配 → 全乱码。
只设 `errors="replace"`。见 [`AGENTS.md` R2](../AGENTS.md#r2--不许把-stdout-强行改成-utf-8-)。

### 5.4 双击启动：`.lnk` 指向 `pythonw.exe`，而不是 `.bat` / `.pyw`

| 方案 | 优点 | 缺点 | 结论 |
| --- | --- | --- | --- |
| `.bat` 包一层 | 简单 | 违反 R1；还会闪一个黑框 | ❌ |
| `.pyw` 文件 | 纯 Python，天然无黑框 | **依赖文件关联**；本机 `HKCR\.pyw` 无默认值，双击只会弹「用什么打开」 | ❌ |
| Windows 快捷方式 `.lnk` | exe / 参数 / 工作目录 / 图标全部写死，不依赖关联；可自定义图标 | 含绝对路径，换机器要重建 | ✅ 采用 |

用 `pythonw.exe` 是为了**没有黑框控制台**。代价是出错时没有任何提示，
所以：`console.ensure_streams()` 先兜住 `sys.stdout is None`
（见 [`AGENTS.md` R22](../AGENTS.md#r22--pythonw-下-sysstdout-是-none必须先补丢弃流-2026-09-28)），
并且 `main()` 在开不了图形界面且没有控制台时会弹系统对话框而不是 `print`。

**创建 `.lnk` 的方式**：`shortcut.py` 用 ctypes 直接调 COM 的
`IShellLinkW` + `IPersistFile`，因此**不需要 pywin32**，也不落 `.ps1`/`.vbs`。
只在 `uv run speedup shortcut` 时跑一次。

**遗留代价**：`.lnk` 里是绝对路径，所以不入 git（`.gitignore` 里 `*.lnk`）；
删过 `.venv` 或换机器后要重跑一次命令重建。

### 5.5 归档区放在 `src/` 外面

`archive/` 是「留档」，不是「模块」。放在 `src/speedup/` 里会被打进 wheel、
被 pytest 收集规则扫到，徒增噪音。

### 5.6 表格排版自己实现，不用 `tabulate`

需求只有一个：**中英混排按显示宽度对齐**。为此引入第三方依赖不划算，
何况 `render_table` 返回字符串（便于断言）这个特性现成库也不给。

---

## 6. 想扩展时从哪下手

| 想做的事 | 改哪 |
| --- | --- |
| 加一个材料识别规则 | `material.py` 的 `STRONG_PATTERNS` / `WEAK_PATTERN` → 加测试 |
| 加一个补料清单列 | `config.py` 的 `SUPPLEMENT_HEADERS` / `SUPPLEMENT_WIDTHS` / `COL_*` |
| 加一个工具 | 见 [`AGENTS.md` §6](../AGENTS.md#6-新增一个工具的标准动作) |
| 给工具加图形界面表单 | 在 `gui.TOOL_FIELDS` 补一份 `Field`（漏了 `TestFieldContract` 会红） |
| 改界面配色 / 字号 | `gui.py` 顶部的颜色常量与 `_setup_fonts()` |
| 改匹配强度 | `config.py` 的 `DEFAULT_MIN_SCORE` 等；或用户传 `--min-score` |
| 换弹窗实现 | 只改 `dialogs.py`，工具层不用动（它们只依赖 `ask_path` 接口） |
| 改快捷方式名字 / 参数 | `shortcut.py` 的 `SHORTCUT_NAME` / `plan_shortcut()` |
| 改图标 | 改 `examples/make_icon.py` 的配色与 `BOLT`，重跑一次 |
| 加输出格式（如导出 CSV） | 在 `excel/` 平级开新子包，别往 `writer.py` 里塞 |
