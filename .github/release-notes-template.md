## ExcelMcp {{VERSION}}

Let your AI assistant create and update real Microsoft Excel workbooks: cells,
formulas, Power Query, PivotTables, charts, and VBA. MCP and CLI offer the same
features. Coding agents can use either; the CLI usually uses fewer AI tokens,
though results vary by task, model, and client.

### What's New
{{CHANGELOG}}

### Get started

**Requires Windows, desktop Excel 2016 or later, and a signed-in interactive
desktop session.**

| Where you work | Setup |
| --- | --- |
| VS Code with GitHub Copilot | Install **Excel MCP Server** from the [VS Code Marketplace](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp). |
| GitHub Copilot plugins | Install `excel-mcp`, `excel-cli`, or both from [Awesome Copilot](https://github.com/github/awesome-copilot). |
| Claude Desktop | Download `excel-mcp-{{VERSION}}.mcpb` below and double-click to install. |
| Another MCP client | [Set up the MCP Server](https://excelmcpserver.dev/installation-mcp-server/). |
| An agent that runs terminal commands | [Set up the CLI](https://excelmcpserver.dev/installation-cli/). |

Node.js 18 or later with npm and `npx` is required for the plugins, Claude
Desktop bundle, and npm installations. The VS Code extension and standalone
executables do not need Node.js or .NET. NuGet tools require .NET 10.

The plugins and VS Code extension include optional report-formatting guidance;
no separate skill download is needed. See the
[installation guide](https://excelmcpserver.dev/installation/) for setup,
updates, and troubleshooting.

### Downloads

| Release asset | What it contains |
| --- | --- |
| `excel-mcp-{{VERSION}}.vsix` | VS Code extension for manual installation. |
| `excel-mcp-{{VERSION}}.mcpb` | Claude Desktop bundle that launches the MCP Server through npm. |
| `ExcelMcp-MCP-Server-{{VERSION}}-windows.zip` | Standalone MCP Server (`mcp-excel.exe`). |
| `ExcelMcp-CLI-{{VERSION}}-windows.zip` | Standalone CLI (`excelcli.exe`). |

Before updating, finish your work and explicitly save and close workbook
sessions. Restart the MCP Server or CLI service to use the updated version;
installing an update does not replace a running process.

### Documentation
- [Website](https://excelmcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- [Changelog](https://github.com/sbroenne/mcp-server-excel/blob/main/CHANGELOG.md)
