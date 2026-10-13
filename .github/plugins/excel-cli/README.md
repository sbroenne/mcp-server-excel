# Excel CLI Plugin

[![npm version](https://img.shields.io/npm/v/%40sbroenne%2Fexcelcli)](https://www.npmjs.com/package/@sbroenne/excelcli)
[![GitHub stars](https://img.shields.io/github/stars/sbroenne/mcp-server-excel?style=flat&label=GitHub%20stars)](https://github.com/sbroenne/mcp-server-excel)

**Let your coding agent create and update Microsoft Excel workbooks from the
terminal: work with cells, formulas, Power Query, PivotTables, charts, and VBA
in the desktop app.**

Ask your agent to enter data, write formulas, refresh Power Query, build
PivotTables and charts, work with DAX, or run VBA. The agent runs `excelcli`
commands and learns them from the built-in help. ExcelMcp operates the actual
Excel application, so Excel handles calculations and workbook features. The
CLI offers the same features as the Excel MCP Server and usually uses fewer AI
tokens.

**Requires Windows, desktop Excel 2016 or later, and Node.js 18 or later with
`npx`. ExcelMcp runs in an interactive desktop session.**

## What you can ask your agent to do

| Request | Result |
| --- | --- |
| "Create a sales tracker with Date, Product, Quantity, Unit Price, and Total columns." | A workbook with a table and formulas. |
| "Create a PivotTable showing revenue by region, then add a chart." | A summary and chart built in Excel. |
| "Refresh the Power Query queries in this workbook and save it." | Updated query results saved in the workbook. |
| "Put this data into a worksheet and format the totals." | Formatted data ready to read. |

The CLI provides 31 feature command categories with 412 operations, including
Power Query, DAX and the Data Model, PivotTables, tables, charts, worksheets,
formulas, VBA, and more. See the
[complete feature reference](https://excelmcpserver.dev/features/).

## Install

Install from [Awesome Copilot](https://github.com/github/awesome-copilot), the
default marketplace in current Copilot clients:

```powershell
copilot plugin install excel-cli@awesome-copilot
```

Or use the direct marketplace:

```powershell
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

Choose one marketplace for this plugin; do not install both copies.

The plugin runs the published npm package through `npx`, so you do not need to
install `excelcli` separately or change your PATH:

```powershell
npx -y @sbroenne/excelcli@latest --help
```

For quoted JSON arguments in Windows PowerShell, the agent can use the plugin's
`bin\start-cli.ps1` wrapper instead. For other coding agents, or to install
`excelcli` as a command on your PATH, follow the
[CLI installation guide](https://excelmcpserver.dev/installation-cli/).

## CLI or MCP?

The CLI and the
[MCP plugin](https://github.com/sbroenne/mcp-server-excel/tree/main/.github/plugins/excel-mcp)
give your agent the same Excel features. You can install either or both; the
difference is only technical:

| Interface | Works with | Context |
|---|---|---|
| CLI (`excelcli`) | Agents that can run terminal commands | Built-in help, read only when needed |
| MCP Server | Any MCP-compatible client | 60 tool schemas loaded into context |

Coding agents can use either one. The CLI usually uses fewer AI tokens because
the agent does not load every tool definition up front. In one comparison using the same task and AI model, the CLI used about 59,000 AI
tokens and MCP used 163,000 (about 64% fewer with CLI). This is one measured
result, not a promise; results vary by task, model, and AI client. See the
[benchmark context](https://github.com/sbroenne/mcp-server-excel/blob/f10926db28c8fc295e15374f4177e3b11ab4bd49/docs/ARCHITECTURE.md).

## More features

- **Workbooks and data:** create, open, save, and close workbooks; edit ranges, tables, and formulas.
- **Analysis:** refresh Power Query, query the Data Model with DAX, and build PivotTables and slicers.
- **Reports:** create and format charts, worksheets, and conditional formatting.
- **Automation:** manage connections, run VBA, and automate Excel windows.
- **Protected workbooks:** detect IRM/AIP protection and open Excel visibly for sign-in.
- **Sessions:** keep a workbook open across commands so the agent can work step by step.

M and DAX code stays on your computer by default. If you set `--format-m-code`
or `--format-dax` to `true`, the corresponding code is sent to
powerqueryformatter.com or daxformatter.com. Only turn it on when you want to
send code to those services.

## Supported agents

- GitHub Copilot CLI
- Other coding agents that can run terminal commands, using the
  [CLI installation guide](https://excelmcpserver.dev/installation-cli/)

## Support

- [Documentation](https://excelmcpserver.dev/)
- [CLI installation guide](https://excelmcpserver.dev/installation-cli/)
- [Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- MIT License
