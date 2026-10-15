# ExcelMcp Plugins

**Let your AI agent create and update real Microsoft Excel workbooks: formulas,
Power Query, PivotTables, charts, and VBA in the desktop app.**

These plugins connect GitHub Copilot CLI or Claude Code to ExcelMcp. ExcelMcp
operates the actual Excel application, so Excel handles calculations and
workbook features.

**Requires Windows, desktop Excel 2016 or later, and Node.js 18 or later with
`npx`. ExcelMcp runs in an interactive desktop session.**

## What you can ask your agent to do

| Request | Result |
| --- | --- |
| "Create a sales tracker with Date, Product, Quantity, Unit Price, and Total columns." | A workbook with a table and formulas. |
| "Create a PivotTable showing revenue by region, then add a chart." | A summary and chart built in Excel. |
| "Refresh the Power Query queries in this workbook and save it." | Updated query results saved in the workbook. |
| "Put this data into a worksheet and format the totals." | Formatted data ready to read. |

You can also work with DAX and the Data Model, connections, slicers, conditional
formatting, and VBA. See the
[complete feature reference](https://excelmcpserver.dev/features/).

## Choose a plugin

| Plugin | How your agent works with Excel | Best fit |
| --- | --- | --- |
| [excel-cli](https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-cli) | Runs terminal commands and reads built-in help when needed. | Coding agents that can run commands; usually uses fewer AI tokens. |
| [excel-mcp](https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-mcp) | Uses Excel tools provided by the MCP Server. | Assistants using MCP tools. |

Both plugins offer the same Excel features; you can install either or both.
The underlying CLI and MCP Server also work outside these plugins; see the
[CLI installation guide](https://excelmcpserver.dev/installation-cli/) or
[MCP Server installation guide](https://excelmcpserver.dev/installation-mcp-server/).

M and DAX code stays on your computer by default. Optional formatting sends the
corresponding code to powerqueryformatter.com or daxformatter.com. Only enable
it when you want to send code to those services; the individual plugin READMEs
describe the options.

## Source and support

- [Documentation](https://excelmcpserver.dev/)
- [Source repository](https://github.com/sbroenne/mcp-server-excel)
- [Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- MIT License

For contributors: this repository is generated publication output. Make changes
in the source repository, not here; publication overwrites unsynchronized edits.
Follow the
[plugin maintenance guide](https://github.com/sbroenne/mcp-server-excel/blob/main/.github/workflows/docs/publish-plugins-setup.md#maintenance-and-updates).

## Install

The plugins are listed in
[Awesome Copilot](https://github.com/github/awesome-copilot), the default
marketplace in current Copilot clients:

```powershell
copilot plugin install excel-mcp@awesome-copilot
copilot plugin install excel-cli@awesome-copilot
```

Alternatively, register this direct marketplace:

```powershell
# Register this marketplace
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install one or both plugins
copilot plugin install excel-mcp@mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

Choose one marketplace per plugin; do not install duplicate copies. Existing
direct-marketplace installations do not need to move.

Both plugins use the public npm packages through `npx`. Node.js 18 or later is
required.

### Claude Code

```powershell
# Register this marketplace
claude plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install one or both plugins
claude plugin install excel-mcp@mcp-server-excel-plugins
claude plugin install excel-cli@mcp-server-excel-plugins
```

Inside a Claude Code session, the same steps are `/plugin marketplace add sbroenne/mcp-server-excel-plugins`
and `/plugin install excel-mcp@mcp-server-excel-plugins`.

For Claude Desktop, use the one-click `.mcpb` bundle from the
[latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest) instead.
