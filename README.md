# mechtools

Some DIY tools for mechanical engineers. Each tool lives in its own folder with
its own README, dependencies and tests — clone once, use what you need.

## Tools

| Tool | What it does | Docs |
| --- | --- | --- |
| **speedup** | Small scripts that cut the paperwork around outsourced part machining: scan a folder of STEP models, read the material off the matching PDF drawings, produce a material-request spreadsheet, then collect the matching models and drawings next to it. Ships with a desktop GUI. | [English](speedup/README.md) · [中文](speedup/README.zh-CN.md) |

More tools will be added here as they are cleaned up for release.

## Layout

```
mechtools/
├── LICENSE
├── README.md          ← you are here
└── speedup/           ← first tool
    ├── README.md
    ├── README.zh-CN.md
    └── src/speedup/
```

## License

Apache-2.0 — see [LICENSE](LICENSE).
