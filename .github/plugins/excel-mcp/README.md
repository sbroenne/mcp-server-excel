# Excel MCP Plugin

**Let your AI assistant create and update Microsoft Excel workbooks using MCP
tools.**

ExcelMcp operates the actual Excel application, so Excel handles calculations
and workbook features.

## Requirements

- GitHub Copilot CLI or Claude Code
- Windows with desktop Excel 2016 or later
- Node.js 18 or later with `npx`

ExcelMcp runs in an interactive desktop session.

## What you can ask your assistant to do

| Request | Result |
| --- | --- |
| "Create a sales tracker with Date, Product, Quantity, Unit Price, and Total columns." | A workbook with a table and formulas. |
| "Create a PivotTable showing revenue by region, then add a chart." | A summary and chart built in Excel. |
| "Refresh the Power Query queries in this workbook and save it." | Updated query results saved in the workbook. |
| "Show me Excel while you work." | Excel is visible as the assistant makes changes. |

## Features

The server provides 63 MCP tools across 33 feature areas, with 429 operations.

- **Workbooks and data:** create or open workbooks, edit ranges and tables, and work with formulas.
- **Analysis:** refresh Power Query, query the Data Model with DAX, and build PivotTables and slicers.
- **Reports:** create and format charts, worksheets, and conditional formatting.
- **Automation:** manage connections, run VBA, and automate Excel windows.
- **Protected workbooks:** detect IRM/AIP protection and open Excel visibly for sign-in.

See the [complete feature reference](https://excelmcpserver.dev/features/).

## Privacy

M and DAX code stays on your computer by default. If you enable
`formatMCode=true` or `formatDax=true`, the corresponding code is sent to
powerqueryformatter.com or daxformatter.com for formatting. Only enable these
options when you want to send code to those services.

## Support

- [Documentation](https://excelmcpserver.dev/)
- [Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- MIT License

## Install

Install from [Awesome Copilot](https://github.com/github/awesome-copilot), the
default marketplace in current Copilot clients:

```powershell
copilot plugin install excel-mcp@awesome-copilot
```

Or use the direct marketplace:

```powershell
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
copilot plugin install excel-mcp@mcp-server-excel-plugins
```

Choose one marketplace for this plugin; do not install both copies.

For Claude Code, use the direct marketplace:

```powershell
claude plugin marketplace add sbroenne/mcp-server-excel-plugins
claude plugin install excel-mcp@mcp-server-excel-plugins
```

The plugin includes the MCP server setup; no manual server configuration is
needed. For Claude Desktop or another MCP-compatible client without this plugin,
follow the
[MCP Server installation guide](https://excelmcpserver.dev/installation-mcp-server/).
