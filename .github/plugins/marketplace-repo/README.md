# ExcelMcp Copilot CLI Plugins

Windows-only GitHub Copilot CLI plugins for ExcelMcp.

This repository is the publish target for plugin artifacts from [`sbroenne/mcp-server-excel`](https://github.com/sbroenne/mcp-server-excel).

## Plugins

- **excel-mcp** — MCP server plugin for conversational Excel automation
- **excel-cli** — CLI plugin for scripting and coding-agent workflows

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

Both plugins publish wrapper/bootstrap assets plus skills. On first use they fetch the newest self-contained Windows runtime from the main `sbroenne/mcp-server-excel` GitHub Releases feed, then reuse it for the rest of the chat session.

## Notes

- **Windows only** — ExcelMcp depends on Microsoft Excel COM automation.
- **excel-mcp** includes MCP configuration plus plugin-local bootstrap helpers for the ExcelMcp MCP runtime.
- **excel-cli** includes plugin-local bootstrap helpers for the Excel CLI runtime; separate PATH installation is optional, not required for plugin use.

## Source and Support

- Source repo: [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel)
- Issues: [sbroenne/mcp-server-excel/issues](https://github.com/sbroenne/mcp-server-excel/issues)
- Plugin docs: [`plugins/excel-mcp/README.md`](./plugins/excel-mcp/README.md), [`plugins/excel-cli/README.md`](./plugins/excel-cli/README.md)

## License

MIT
