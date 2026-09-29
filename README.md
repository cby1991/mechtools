# mechtools

Some DIY tools for mechanical engineers. Each tool lives in its own folder with
its own README, dependencies and tests — clone once, use what you need.

## Tools

| Tool | What it does | Docs |
| --- | --- | --- |
| **speedup** | Small scripts that cut the paperwork around outsourced part machining: scan a folder of STEP models, read the material off the matching PDF drawings, produce a material-request spreadsheet, then collect the matching models and drawings next to it. Ships with a desktop GUI. | [English](speedup/README.md) · [中文](speedup/README.zh-CN.md) |
| **swauto** | A SolidWorks VBA macro. One click exports models to STEP and their drawings to PDF. Expanding an assembly automatically skips standard parts (screws, nuts, washers) and lets you choose which of hidden / suppressed / envelope components to include. | [中文](swauto/README.md) |

More tools will be added here as they are cleaned up for release.

## Layout

```
mechtools/
├── LICENSE
├── README.md          ← you are here
├── speedup/           ← pure Python, desktop GUI
│   ├── README.md
│   ├── README.zh-CN.md
│   └── src/speedup/
└── swauto/            ← SolidWorks VBA macro
    ├── README.md
    ├── src/           ← 8 source segments (UTF-8)
    ├── dist/          ← importable single-file macro (GBK + CRLF)
    └── tests/         ← pytest, runs without SolidWorks
```

Each tool keeps its own stack and its own toolchain. `speedup` is pure Python;
`swauto` is VBA plus a Python build/check script. They share nothing but the repo.

## License

Apache-2.0 — see [LICENSE](LICENSE).
