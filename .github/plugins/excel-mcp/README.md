# Excel MCP Plugin

[![npm version](https://img.shields.io/npm/v/%40sbroenne%2Fmcp-server-excel)](https://www.npmjs.com/package/@sbroenne/mcp-server-excel)
[![GitHub stars](https://img.shields.io/github/stars/sbroenne/mcp-server-excel?style=flat&label=GitHub%20stars)](https://github.com/sbroenne/mcp-server-excel)

**Let your AI assistant create and update Microsoft Excel workbooks: work with
cells, formulas, Power Query, PivotTables, charts, and VBA in the desktop app.**

Ask your assistant to enter data, write formulas, refresh Power Query, build
PivotTables and charts, work with DAX, or run VBA. ExcelMcp operates the actual
Excel application, so Excel handles calculations and workbook features.

**Requires Windows, desktop Excel 2016 or later, an MCP-compatible client, and
Node.js 18 or later with `npx`. ExcelMcp runs in an interactive desktop session.**

## What you can ask your assistant to do

| Request | Result |
|---|---|
| "Create a sales tracker with Date, Product, Quantity, Unit Price, and Total columns." | A workbook with a table and formulas. |
| "Create a PivotTable showing revenue by region, then add a chart." | A summary and chart built in Excel. |
| "Refresh the Power Query queries in this workbook and save it." | Updated query results saved in the workbook. |
| "Show me Excel while you work." | Excel is visible as the assistant makes changes. |

The server provides 60 MCP tools across 31 feature areas, with 408 operations,
including Power Query, DAX and the Data Model, PivotTables, tables, charts,
worksheets, formulas, VBA, and more. See the
[complete feature reference](https://excelmcpserver.dev/features/).

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

The plugin includes its MCP launch configuration and runs the published npm
package through `npx`:

```powershell
npx -y @sbroenne/mcp-server-excel@latest
```

You do not need to add a separate MCP server entry to
`~/.copilot/mcp-config.json`. For another MCP-compatible client, follow the
[MCP Server installation guide](https://excelmcpserver.dev/installation-mcp-server/).

## MCP or CLI?

Use the MCP Server when you want your assistant to find and use Excel tools in a
conversation. Use the
[CLI plugin](https://github.com/sbroenne/mcp-server-excel/tree/main/.github/plugins/excel-cli)
primarily for coding agents using terminal commands. The CLI can also be used
from scripts.

In one comparison using the same task and AI model, the CLI used about 59,000 AI
tokens and MCP used 163,000 (about 64% fewer with CLI). This is one measured
result, not a promise; results vary by task, model, and AI client. See the
[benchmark context](https://github.com/sbroenne/mcp-server-excel/blob/f10926db28c8fc295e15374f4177e3b11ab4bd49/docs/ARCHITECTURE.md).

## More features

- **Workbooks and data:** create or open workbooks, edit ranges and tables, and work with formulas.
- **Analysis:** refresh Power Query, query the Data Model with DAX, and build PivotTables and slicers.
- **Reports:** create and format charts, worksheets, and conditional formatting.
- **Automation:** manage connections, run VBA, and automate Excel windows.
- **Protected workbooks:** detect IRM/AIP protection and open Excel visibly for sign-in.

M and DAX code stays on your computer by default. If you enable
`formatMCode=true` or `formatDax=true`, the corresponding code is sent to
powerqueryformatter.com or daxformatter.com for formatting. Only enable these
options when you want to send code to those services.

## Supported clients

- GitHub Copilot
- Claude Desktop
- Cursor
- Cline
- Windsurf
- Other MCP-compatible clients

## Support

- [Documentation](https://excelmcpserver.dev/)
- [MCP Server installation guide](https://excelmcpserver.dev/installation-mcp-server/)
- [Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- MIT License
