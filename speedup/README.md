# speedup — Workshop Speed-up Toolkit

**[中文说明 / Chinese](README.zh-CN.md)**

Small tools that automate the paperwork around outsourced part machining: scan a
folder of STEP models, read the material off the matching PDF drawings, produce a
material-request spreadsheet — then find those models and drawings again and
collect them next to the spreadsheet.

Pure Python. No `.bat`, no PowerShell. Ships with a desktop GUI, a console menu,
subcommands, and a test suite.

## What it does

| Tool | Subcommand | What it does |
| --- | --- | --- |
| Generate material list | `gbbuild` | Scans a folder for `.stp` / `.step`, finds the **same-named PDF drawing** for each model, reads the material from the drawing's title block, and writes an Excel list. |
| Match & copy files | `gbcopy` | Reads the part names from that Excel, finds the matching STEP + PDF in a search folder, and copies them next to the spreadsheet. **Never overwrites an existing file.** |

## Requirements

| | |
| --- | --- |
| OS | Windows 10/11. The core logic is portable, but the GUI, the shortcut and "open with default app" are Windows-oriented. |
| Python | **3.12** — managed by uv, you don't install it yourself. |
| uv | ≥ 0.5 |

> **Why pin 3.12?** The GUI needs `tkinter`, and uv's Python 3.12 distribution
> bundles it (Tk 8.6). The standalone 3.13 build may not.

## Install

```bash
git clone <your-repo-url> speedup
cd speedup
uv sync
uv run speedup doctor      # sanity check
uv run speedup shortcut    # create the double-click entry point
```

## Usage

### Double-click (daily driver)

`uv run speedup shortcut` creates `启动 speedup.lnk` in the project root
(add `--desktop` to place it on the Desktop instead). Double-click it to open the
GUI — **no console window**, because it runs under `pythonw.exe`.

### GUI

```bash
uv run speedup          # opens the GUI
uv run speedup gui      # same, but fails loudly instead of falling back
```

One page per tool: pick paths with *Browse…*, set options, hit **Run**, watch the
log. Tasks run on a background thread, so the window never freezes.

### Console menu

```bash
uv run speedup menu
```

### Subcommands

```bash
uv run speedup gbbuild "D:\example\batch01"
uv run speedup gbcopy "list.xlsx" "D:\example\models" --dry-run
```

| Other command | |
| --- | --- |
| `uv run speedup doctor` | environment check |
| `uv run speedup archive` | list the local archive of older scripts |
| `uv run speedup shortcut` | (re)create the double-click shortcut |

## Notes

- **Nothing is overwritten.** `gbcopy` skips any file that already exists in the
  target folder and lists it in the report — that file may have been edited by hand.
- **One worksheet only.** A workbook often contains extra sheets that also have a
  name column (a reference sheet, for instance). `gbcopy` reads exactly one: the
  sheet named `补料清单` if present, otherwise the first candidate. It reports
  which sheets it ignored; override with `--sheet <name>`.
- **`archive/` is local-only.** The repository deliberately does not track it —
  it holds historical one-off scripts containing internal paths and project names.
  `uv run speedup archive` still works on a machine that has the folder.

## Tests

```bash
uv run pytest                                            # 343 passed, 1 skipped
uv run pytest --cov=speedup --cov-report=term-missing     # coverage
```

Tests build their data under `tmp_path` and never touch real folders.
The GUI tests don't open windows — they drive a hidden Tk root directly.

## Layout

```
speedup/
├── src/speedup/
│   ├── cli.py           # entry point: GUI / menu / subcommands
│   ├── gui.py           # tkinter GUI — builds argv, nothing else
│   ├── shortcut.py      # creates the .lnk (ctypes + COM, no pywin32)
│   ├── config.py        # every constant lives here
│   ├── console.py       # display-width-aware table output
│   ├── dialogs.py       # tkinter file/folder pickers
│   ├── material.py      # material detection rules (pure functions)
│   ├── matching.py      # match scoring + copy planning (pure functions)
│   ├── excel/           # xlsx reader & writer
│   └── tools/           # one module per tool
├── tests/               # 12 files, 344 cases
└── docs/ARCHITECTURE.md # layering and the trade-offs behind it
```

## Design rules

[`AGENTS.md`](AGENTS.md) documents the hard constraints and why they exist
(written in Chinese). The three that shape everything:

1. **No `.bat` / `.cmd` / `.ps1` / `.vbs`.** Interaction goes through tkinter,
   orchestration through subcommands. Checked by
   `git ls-files | grep -E '\.(bat|cmd|ps1|vbs)$'` returning nothing.
2. **Never force `sys.stdout` to UTF-8.** Chinese Windows consoles default to
   cp936; changing the Python-side encoding turns everything into mojibake.
3. **The GUI contains no business logic.** It only collects values into an argv
   list and hands it to the same `run()` the CLI calls — otherwise "works in the
   GUI, broken on the command line" becomes possible.
