# Excel CLI Plugin

**Run commands and scripts to work with real Microsoft Excel workbooks, formulas,
Power Query, PivotTables, charts, and VBA.**

Use `excelcli` to enter data, write formulas, refresh Power Query, build PivotTables
and charts, or automate repeatable workbook tasks. ExcelMcp operates the actual
Excel application, so Excel handles calculations and workbook features.

**Requires Windows, desktop Excel 2016 or later, and Node.js 18 or later with
`npx`. ExcelMcp runs in an interactive desktop session.**

## What you can ask your agent to do

- "Create a sales workbook with a table, a revenue formula, and a chart."
- "Refresh the Power Query queries in this workbook and save it."
- "Add a PivotTable showing total sales by region."
- "Put this data into a worksheet and format the totals."

The CLI covers Power Query, DAX and the Data Model, PivotTables, tables, charts,
formulas, VBA, worksheets, workbook files, and more. See the
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

## Run the CLI

The plugin does not bundle `excelcli.exe` or add the `excelcli` command to your
PATH. Its wrapper launches the current npm package with `npx`:

```powershell
npx -y @sbroenne/excelcli@latest --help
```

No global helper, PATH change, or separate runtime installation is needed.
For quoted JSON arguments in Windows PowerShell, use the plugin's
`bin\start-cli.ps1` wrapper:

```powershell
& "C:\Path\To\Installed\excel-cli\bin\start-cli.ps1" --help
```

Use the installed plugin directory in place of the example path. The examples
below use `excelcli` for readability; if you have not installed the standalone
CLI, substitute `npx -y @sbroenne/excelcli@latest` or use the wrapper.

```powershell
# Create a workbook
excelcli -q session create C:\Reports\Sales.xlsx

# Write headers and data
excelcli -q range set-values --session <id> --sheet Sheet1 `
  --range A1:C1 `
  --values '[["Date","Product","Revenue"]]'
excelcli -q range set-values --session <id> --sheet Sheet1 `
  --range A2:C3 `
  --values '[["2024-01-15","Widget",1500],["2024-01-16","Gadget",2300]]'

# Create a table, then save and close
excelcli -q table create --session <id> --sheet Sheet1 `
  --table-name SalesData --range A1:C3
excelcli -q session close --session <id> --save
```

### Optional standalone installation

If you prefer to install the CLI separately from the plugin, choose one:

**Windows executable**

1. Download `ExcelMcp-CLI-{version}-windows.zip` from
   [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest).
2. Extract `excelcli.exe` to a permanent folder, such as `C:\Tools\ExcelMcp\`.
3. Add that folder to your PATH.

**.NET global tool**

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
# Requires the .NET 10 Runtime
```

## CLI or MCP Server?

Both interfaces automate the same Excel operations. The CLI is a good fit for
scripts and coding-agent workflows; the MCP Server is a good fit for
conversational tool discovery.

In one comparison using the same task and AI model, the CLI used about 59,000 AI
tokens and MCP used 163,000 (about 64% fewer with CLI). This is one measured
result, not a promise; results vary by task, model, and AI client. See the
[benchmark context](https://github.com/sbroenne/mcp-server-excel/blob/f10926db28c8fc295e15374f4177e3b11ab4bd49/docs/ARCHITECTURE.md).

## More features

- **Workbooks and files:** create, open, save, close, and export workbooks.
- **Data and formulas:** read and write ranges, tables, formulas, and named ranges.
- **Analysis:** use Power Query, DAX, the Data Model, PivotTables, and slicers.
- **Reports:** create and format charts, worksheets, and conditional formatting.
- **Automation:** manage connections, run VBA, and automate Excel windows.
- **Protected workbooks:** detect IRM/AIP protection and open Excel visibly for sign-in.

## Support

- [Documentation](https://excelmcpserver.dev/)
- [Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- MIT License
