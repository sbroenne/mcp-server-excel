# ExcelMcp Copilot CLI Plugins

Windows-only GitHub Copilot CLI plugins for ExcelMcp.

This repository is the publish target for plugin artifacts from [`sbroenne/mcp-server-excel`](https://github.com/sbroenne/mcp-server-excel).

> [!WARNING]
> This repository is generated publication output. Do not edit it directly:
> publication overwrites unsynchronized changes. Update the canonical files in
> the [source repository](https://github.com/sbroenne/mcp-server-excel) and
> follow its [plugin publication guide](https://github.com/sbroenne/mcp-server-excel/blob/main/.github/workflows/docs/publish-plugins-setup.md#maintenance-and-updates).

## Plugins

- **excel-mcp** — MCP server plugin for conversational Excel automation
- **excel-cli** — CLI plugin for scripting and coding-agent workflows

## Repository Layout

```text
.github/plugin/marketplace.json
plugins/
├── excel-mcp/
│   ├── plugin.json
│   ├── mcp.json
│   └── skills/excel-mcp-report-formatting/SKILL.md
└── excel-cli/
    ├── plugin.json
    └── skills/excel-cli-report-formatting/SKILL.md
```

The canonical marketplace manifest lives at `.github/plugin/marketplace.json`. The `plugins/` directory contains Agent Plugins 1.0 packages generated from source-owned templates by the source repo's `publish-plugins.yml` workflow.

## Install

```powershell
# Register this marketplace
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install one or both plugins
copilot plugin install excel-mcp@mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

Both plugins use the public npm packages through `npx`. Node.js 18 or later is
required.

## Notes

- **Windows only** — ExcelMcp depends on Microsoft Excel COM automation.
- **excel-mcp** includes portable root `mcp.json` configuration that launches `@sbroenne/mcp-server-excel`.
- **excel-cli** includes an argument-safe npx wrapper for `@sbroenne/excelcli`; separate PATH installation is optional.
- Both root `plugin.json` manifests target `https://agent-plugins.org/schemas/1.0.0/plugin.schema.json`; skills are discovered from the fixed `skills/` directory.

## Source and Support

- Source repo: [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel)
- Issues: [sbroenne/mcp-server-excel/issues](https://github.com/sbroenne/mcp-server-excel/issues)
- Plugin docs: [excel-mcp](https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-mcp), [excel-cli](https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-cli)

## License

MIT
