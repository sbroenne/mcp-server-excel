# ExcelMcp Copilot CLI Plugins

Windows-only GitHub Copilot CLI plugins for ExcelMcp.

This repository is the publish target for plugin artifacts from [`sbroenne/mcp-server-excel`](https://github.com/sbroenne/mcp-server-excel).

## Plugins

- **excel-mcp** — MCP server plugin for conversational Excel automation
- **excel-cli** — CLI skill plugin for scripting and coding-agent workflows

## Repository Layout

```text
.github/plugin/marketplace.json
plugins/
├── excel-mcp/
└── excel-cli/
```

The canonical marketplace manifest lives at `.github/plugin/marketplace.json`. The `plugins/` directory contains the distributable plugin bundles published by the source repo's `publish-plugins.yml` workflow.

## Install

```powershell
# Register this marketplace
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install one or both plugins
copilot plugin install excel-mcp@mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

The `excel-cli` plugin is skill-only. Install `excelcli` separately from the main release ZIP or NuGet tool if you want the command on PATH.

## Notes

- **Windows only** — ExcelMcp depends on Microsoft Excel COM automation.
- **excel-mcp** includes MCP configuration plus helper scripts for the ExcelMcp MCP server.
- **excel-cli** provides the skill package only; install `excelcli` separately from the source repo releases if needed.

## Source and Support

- Source repo: [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel)
- Issues: [sbroenne/mcp-server-excel/issues](https://github.com/sbroenne/mcp-server-excel/issues)
- Plugin docs: [`plugins/excel-mcp/README.md`](./plugins/excel-mcp/README.md), [`plugins/excel-cli/README.md`](./plugins/excel-cli/README.md)

## License

MIT
