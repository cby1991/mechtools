# speedup —— 车间提速工具集

**[English](README.md)**

把外发加工那些跑腿的纸面活儿自动化：扫一个装满 STEP 模型的文件夹，从同名 PDF
图纸里读出材料和表面处理要求，生成一份采购清单 Excel；之后再按这份清单把对应的
模型和图纸找回来，收拢到清单旁边。

**纯 Python 实现，不依赖 `.bat` / PowerShell。** 带桌面图形界面、文字菜单、子命令和测试。

## 有什么工具

| 工具 | 子命令 | 干什么 |
| --- | --- | --- |
| 生成采购清单 | `gbbuild` | 扫文件夹里的 `.stp` / `.step`，为每个模型找**同名的 PDF 图纸**，从标题栏读出材料、从「技术要求」栏读出表面处理要求（写进备注列），生成 Excel 清单。加 `--with-image` 还能给每个零件渲染一张**轴测图**填进「图片」列（需另装 `speedup[images]`）。 |
| 采购文件匹配拷贝 | `gbcopy` | 读那份 Excel 里的零件名，在搜索文件夹里找回对应的 STEP + PDF，拷到清单旁边。**同名文件绝不覆盖。** |

## 环境要求

| | |
| --- | --- |
| 操作系统 | Windows 10/11。核心逻辑跨平台，但图形界面、快捷方式、「用默认程序打开」都是按 Windows 做的。 |
| Python | **3.12** —— 由 uv 自动管理，不用自己装。 |
| uv | ≥ 0.5 |

> **为什么锁 3.12？** 图形界面要 `tkinter`，而 uv 的 Python 3.12 分发包自带
> （Tk 8.6）。3.13 的独立分发包不一定带。

## 安装

```bash
git clone https://github.com/cby1991/mechtools.git
cd mechtools/speedup
uv sync
uv run speedup doctor      # 环境体检
uv run speedup shortcut    # 生成双击启动的图标
```

## 怎么用

### 双击启动（日常最省事）

`uv run speedup shortcut` 会在项目根目录生成 `启动 speedup.lnk`（加 `--desktop`
就放桌面）。**双击它**即可打开图形界面，而且**没有黑框控制台** —— 因为它跑的是
`pythonw.exe`。

### 图形界面

```bash
uv run speedup          # 打开图形界面
uv run speedup gui      # 同上，但开不了窗口时明确报错，不静默降级
```

一个工具一页：点「浏览…」选路径 → 设参数 → 点「Run」→ 看日志。
任务跑在后台线程，窗口不会假死。

### 文字菜单

```bash
uv run speedup menu
```

### 子命令

```bash
uv run speedup gbbuild "D:\example\batch01"
uv run speedup gbcopy "清单.xlsx" "D:\example\models" --dry-run
```

| 其他命令 | |
| --- | --- |
| `uv run speedup doctor` | 环境体检 |
| `uv run speedup archive` | 列出本机留档的历史脚本 |
| `uv run speedup shortcut` | 生成／重建双击启动的快捷方式 |

## 几条要注意的

- **不覆盖任何东西。** `gbcopy` 遇到目标文件夹里已有同名文件会跳过并列进报告 ——
  那有可能是人工改过的。
- **只读一张工作表。** 工作簿里常常还有别的表也带名称列（比如"参考表"）。
  `gbcopy` 只认一张：优先名为 `采购清单` 的那张，没有就取第一张。
  它会**告诉你忽略了哪几张**；要换表用 `--sheet 表名`。
- **`archive/` 只在本地。** 仓库**刻意不跟踪**这个目录 —— 里面是历史一次性脚本，
  含内部路径和项目名。本机有目录时 `uv run speedup archive` 照常可用。

## 跑测试

```bash
uv run pytest                                            # 343 passed, 1 skipped
uv run pytest --cov=speedup --cov-report=term-missing     # 覆盖率
```

测试数据全部造在 `tmp_path` 下，**不碰你的真实目录**。
图形界面的测试**不弹窗口** —— 用一个隐藏的 Tk 根窗口直接调方法。

## 目录结构

```
speedup/
├── src/speedup/
│   ├── cli.py           # 入口：图形界面 / 菜单 / 子命令
│   ├── gui.py           # tkinter 图形界面 —— 只负责拼 argv
│   ├── shortcut.py      # 生成 .lnk（ctypes 调 COM，不用 pywin32）
│   ├── config.py        # 所有常量集中在这里
│   ├── console.py       # 按显示宽度对齐的表格输出
│   ├── dialogs.py       # tkinter 文件/文件夹选择框
│   ├── material.py      # 材料识别规则（纯函数）
│   ├── matching.py      # 匹配打分 + 拷贝规划（纯函数）
│   ├── excel/           # xlsx 读 / 写
│   └── tools/           # 一个工具一个模块
├── tests/               # 11 个文件，344 个用例
└── docs/ARCHITECTURE.md # 分层设计与取舍记录
```

## 设计准则

硬约束和它们背后的理由记在 [`AGENTS.md`](AGENTS.md)（中文）。
最关键的三条：

1. **不许有 `.bat` / `.cmd` / `.ps1` / `.vbs`。** 交互走 tkinter，编排走子命令。
   验收：`git ls-files | grep -E '\.(bat|cmd|ps1|vbs)$'` 无输出。
2. **不许把 `sys.stdout` 改成 UTF-8。** 中文 Windows 控制台默认 cp936，
   改 Python 侧编码会全变乱码。
3. **图形界面里不许有业务逻辑。** 它只把控件值拼成 argv，交给命令行用的同一个
   `run()` —— 否则就会出现「界面上能跑、命令行不行」。
