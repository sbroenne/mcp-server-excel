# ExcelMcp - Real Excel Automation for VS Code

<a href="https://github.com/sbroenne/mcp-server-excel"><img src="https://img.shields.io/github/stars/sbroenne/mcp-server-excel?style=flat&label=GitHub%20Stars" alt="GitHub stars" width="112" height="20"></a>
<a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT" width="81" height="20"></a>

**Automate real Microsoft Excel with GitHub Copilot.**

Refresh Power Query, calculate formulas, build PivotTables and charts, work with
DAX measures, or run VBA directly from Copilot Chat.

**Requires Windows, desktop Excel 2016 or later, VS Code 1.125 or later, and
GitHub Copilot chat with tool support.**

The server and Excel guidance are included: **no separate .NET, Node.js, CLI,
or skill installation is needed**.

**Excel does the work.** Calculations, Power Query refreshes, and Data Model
operations run in the actual Excel application. Excel opens and saves your
workbook itself, rather than a file-parser library rewriting it. Requested
edits can still change the workbook's data, formatting, or features.

## See It in Action

<a href="https://excelmcpserver.dev/use-cases/"><img src="https://excelmcpserver.dev/assets/images/excel-demo-table-chart.png" alt="A sales table, regional summary, and column chart created in real Microsoft Excel" width="680" height="400"></a>

A sales table, regional summary, and chart created in Excel. Ask Copilot:
*"Show me Excel while you work"* to watch changes live.

<a href="https://youtu.be/wbw3-hPcE2o"><img src="https://img.youtube.com/vi/wbw3-hPcE2o/maxresdefault.jpg" alt="Watch the two-minute Excel MCP Server intro video" width="384" height="216"></a>

[Watch the intro video (2 min)](https://youtu.be/wbw3-hPcE2o)

## Quick Start

1. **Install the extension** on your Windows desktop with Excel installed.
2. **Open Copilot Chat with tool support.**
3. **Give Copilot a local workbook path or ask it to create a workbook.** With
   VS Code's default settings, the bundled **excel-mcp** server starts
   automatically when your request needs Excel tools. Approve server or tool
   use if prompted. Try:
   "Create a new workbook on my Desktop with sample sales data and a column
   chart. Save it as SalesDemo.xlsx and show me Excel while you work."

Close a workbook in Excel before asking Copilot to open it: ExcelMcp needs
exclusive access while automating it.

## What You Can Ask Copilot

| Ask Copilot | Result |
|---|---|
| "Refresh every Power Query query in my sales workbook, then save it." | An updated workbook with fresh query results. [Refresh guide](https://excelmcpserver.dev/guides/refresh-power-query/) |
| "Create a PivotTable showing revenue by region, add a column chart, and save the workbook." | A summary and chart built in Excel. [PivotTable guide](https://excelmcpserver.dev/guides/automate-pivottables/) |
| "Export the Power Query M code and VBA modules from my workbook to source files." | Query and macro source files you can review and keep in version control. [Power Query reference](https://excelmcpserver.dev/reference/powerquery/) |

## Key Features

Excel MCP Server (excel-mcp) provides **60 MCP tools across 31 feature areas, with 400 operations**:

- **Power Query & M code** - Import data, create and edit queries, refresh results, and choose load destinations.
- **Power Pivot & DAX** - Build Data Models, create measures, and manage relationships.
- **PivotTables, charts & slicers** - Build and format reports from tables, ranges, or the Data Model.
- **Data & workbooks** - Read and write data, formulas, and formatting; filter and sort tables; manage worksheets, named ranges, and connections.
- **VBA macros** - Inspect, import, update, run, and export VBA modules.
- **Python in Excel** - Write and run `=PY()` formulas using Excel's cloud Python engine.
- **Watch Copilot work** - See Excel side by side with live status feedback.

[See all 60 MCP tools and 400 operations](https://excelmcpserver.dev/features/).
Tool workflows are tested with real AI assistants using
[pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering).

### Optional Report Formatting Included

Copilot can load the bundled
[`excel-mcp-report-formatting` skill](https://excelmcpserver.dev/skills/)
for requested report presentation. Ordinary workbook work uses native MCP
guidance rather than a broad automatically loaded skill.
No separate skill installation or preview setting is required.
Use ordinary natural-language requests; `/skills` opens VS Code's Configure
Skills menu if you want to inspect available skills.

## Requirements

- **Windows x64 or Windows ARM64** with an interactive desktop.
- **Microsoft Excel 2016 or later**, installed and able to open normally.
  Some features, such as Python in Excel, require a supported Excel edition.
- **VS Code 1.125 or later** and GitHub Copilot chat with tool support.

This extension is not for macOS, Linux, browser-only VS Code, Windows services,
or unattended server-side processing. It bundles the MCP server, not `excelcli`.

## Troubleshooting

| Problem | What to Do |
|---|---|
| "Excel is not installed" | Install desktop Excel 2016 or later and check that it opens normally. |
| "VBA access denied" | In Excel, open **File > Options > Trust Center > Trust Center Settings > Macro Settings** and enable **Trust access to the VBA project object model**. This one-time setup is required for VBA operations. |
| Copilot cannot see Excel tools | Run **MCP: List Servers**, choose **excel-mcp**, and start it. Enable Excel tools in chat and accept the server trust prompt. After an update, refresh tools when prompted. |
| "Bundled server is missing or unreadable" | Check security software and file permissions, or reinstall the extension. |
| "Could not check Excel registration" or a registration timeout | Check that Windows PowerShell and desktop Excel open normally. Retry; repair Microsoft Office if the installation is damaged. |

- **Server logs:** run **MCP: List Servers**, choose **excel-mcp**, then
  **Show Output**.
- **Extension setup diagnostics:** open the Output panel and choose
  **ExcelMcp**, or select **Show Setup Output** in a setup error notification.

Startup checks read Excel's registration; they do not start Excel, open a
workbook, or verify that every Excel feature is available.

## Privacy

Excel runs on your Windows desktop. Requested workbook results are returned to
your AI assistant, whose privacy policy applies. Release builds can send
anonymous usage statistics, but those statistics exclude workbook contents,
file names, and paths. Optional remote M/DAX formatting and Python in Excel
use external services only when you request those features.
[Read the privacy policy](https://excelmcpserver.dev/privacy/).

## Guides & Support

- [Refresh Power Query](https://excelmcpserver.dev/guides/refresh-power-query/)
- [Build and update PivotTables](https://excelmcpserver.dev/guides/automate-pivottables/)
- [Query the Data Model with DAX](https://excelmcpserver.dev/guides/query-data-model-with-dax/)
- [Run VBA macros](https://excelmcpserver.dev/guides/run-vba-macros/)

[All user guides](https://excelmcpserver.dev/guides/) (also opened by
**Getting Started**) |
[Complete documentation](https://excelmcpserver.dev/) |
[Report an issue](https://github.com/sbroenne/mcp-server-excel/issues)

MIT License - see [LICENSE](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE).
