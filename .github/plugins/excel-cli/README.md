# Excel CLI Plugin

**Let your coding agent create and update Microsoft Excel workbooks from the
terminal.**

Your agent runs `excelcli` commands and reads built-in help when needed.
ExcelMcp operates the actual Excel application, so Excel handles calculations
and workbook features.

## Requirements

- GitHub Copilot CLI or Claude Code
- Windows with desktop Excel 2016 or later
- Node.js 18 or later with `npx`
- PowerShell 7.4 or later (`pwsh`) when using `bin\start-cli.ps1`;
  Windows PowerShell 5.1 is not supported

ExcelMcp runs in an interactive desktop session.

## What you can ask your agent to do

| Request | Result |
| --- | --- |
| "Create a sales tracker with Date, Product, Quantity, Unit Price, and Total columns." | A workbook with a table and formulas. |
| "Create a PivotTable showing revenue by region, then add a chart." | A summary and chart built in Excel. |
| "Refresh the Power Query queries in this workbook and save it." | Updated query results saved in the workbook. |
| "Put this data into a worksheet and format the totals." | Formatted data ready to read. |

## Features

The CLI provides 33 feature command categories with 431 operations.

- **Workbooks and data:** create, open, save, and close workbooks; edit ranges, tables, and formulas.
- **Analysis:** refresh Power Query, query the Data Model with DAX, and build PivotTables and slicers.
- **Reports:** create and format charts, worksheets, and conditional formatting.
- **Automation:** manage connections, run VBA, and automate Excel windows.
- **Protected workbooks:** detect IRM/AIP protection and open Excel visibly for sign-in.
- **Sessions:** keep a workbook open across commands so the agent can work step by step.

See the [complete feature reference](https://excelmcpserver.dev/features/).

## Privacy

M and DAX code stays on your computer by default. If you set `--format-m-code`
or `--format-dax` to `true`, the corresponding code is sent to
powerqueryformatter.com or daxformatter.com. Only turn it on when you want to
send code to those services.

## Support

- [Documentation](https://excelmcpserver.dev/)
- [Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- MIT License

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

For Claude Code, use the direct marketplace:

```powershell
claude plugin marketplace add sbroenne/mcp-server-excel-plugins
claude plugin install excel-cli@mcp-server-excel-plugins
```

The plugin launches `excelcli` automatically; no separate installation or PATH
change is needed. To use the CLI without this plugin, follow the
[CLI installation guide](https://excelmcpserver.dev/installation-cli/).
