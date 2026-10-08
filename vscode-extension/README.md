# Excel MCP Server - AI-Powered Excel Automation

[![GitHub](https://img.shields.io/badge/GitHub-sbroenne%2Fmcp--server--excel-blue)](https://github.com/sbroenne/mcp-server-excel)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)


**Control Microsoft Excel with AI through GitHub Copilot - just ask in natural language!**

> **Apple Silicon macOS support is experimental beta.** Windows retains the
> complete backend. Power Query, VBA, Data Model/DAX/OLAP, Tables, PivotTables,
> charts, slicers, connections, QueryTables, XML Maps, screenshots, advanced
> visual formatting, and Python result reads are not supported on Mac. See
> [macOS beta limitations](https://github.com/sbroenne/mcp-server-excel/blob/main/specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta)
> before use. Failed or cancelled mutations can partly apply; inspect the
> surviving session before retrying. The feature list and analytics examples
> below describe Windows capabilities, not Mac beta availability.

**MCP Server for Excel** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands. Automate Power Query, DAX measures, VBA macros, PivotTables, Charts, formatting, and data transformations - no Excel programming knowledge required.

**⚡ Powered by the real Excel engine** - ExcelMcp automates the **actual Excel application** through COM on Windows and a capability-gated Apple Events backend on macOS. On Windows, that unlocks what spreadsheets are really for:

- **Runs live Excel operations** - Refresh Power Query to pull and reshape fresh data, recalculate with Excel's own engine, refresh PivotTables and the Data Model, evaluate DAX, and run VBA or Python `=PY()` — the real, *computed results* land right in your workbook.
- **Uses Excel to open and save your files** - No file-parser rewrite of workbook
  contents. Failed or cancelled mutations are not guaranteed to roll back or
  preserve every object unchanged; inspect the surviving session before retrying.

Other tools (openpyxl-based MCP servers and Agent Skills, including Anthropic's `xlsx` skill) read and rewrite the `.xlsx` file directly — which can quietly drop PivotTables, charts, and macros, and can't run Power Query, the Data Model, or DAX at all. Here, Excel does the work. Watch it live: just say *"Show me Excel while you work."*

**💡 Interactive Development** - See results instantly in Excel. Create a query, run it, inspect the output, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing.

## Key features

The Excel MCP Server (excel-mcp) provides **31 specialized tools with 387 operations** for comprehensive Excel automation:

- 🔄 **Power Query & M code** - Create, edit and optimize M code. Import from files, databases and APIs. Refresh queries and manage load destinations.
- 🧮 **Power Pivot & DAX** - Build Data Models, create DAX measures and manage table relationships. Full Power Pivot automation.
- 📊 **PivotTables & charts** - Create PivotTables from ranges, tables or the Data Model. Build charts and PivotCharts with full formatting control.
- 📋 **Tables & ranges** - Read/write data, formulas and formatting. Filter, sort and validate. Manage Excel Tables with structured references.
- 📝 **VBA macros** - View, import, update and execute VBA code. Export modules for version control.
- 📄 **Worksheets & connections** - Manage sheets, named ranges and data connections. Copy and move sheets between workbooks.
- 👁️ **Agent mode** - Watch AI work in Excel in real time — side-by-side view, live status-bar feedback and smart window arrangement, like a pair programmer in a spreadsheet.
- 🐍 **Python in Excel** - Write and run `=PY()` formulas that execute in Excel's cloud Python engine — process worksheet data with pandas, NumPy and more, from your AI assistant.
- 🧪 **LLM-tested quality** - Tool behavior validated with real LLM workflows using [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering), so AI assistants reliably understand and use every operation.

📚 **[See all 31 tools and 387 operations →](https://excelmcpserver.dev/features/)**

### Agent Skills (Bundled)

This extension includes an **Agent Skill** following the [agentskills.io](https://agentskills.io) specification - providing domain-specific guidance for AI assistants:

- **[excel-mcp](https://excelmcpserver.dev/skills/)** - MCP Server tool guidance

The skill is registered automatically through VS Code's `chatSkills`
contribution point. No separate skill installation or preview setting is needed.


## 💬 Example Prompts

**Create & Populate Data (plain ranges on Mac; Tables on Windows):**
- *"Create a new Excel file called SalesTracker.xlsx with a table for Date, Product, Quantity, Unit Price, and Total"*
- *"Put this data in A1:C4 - Name, Age, City / Alice, 30, Seattle / Bob, 25, Portland"*
- *"Add sample data and a formula column for Quantity times Unit Price"*

**Analysis & Visualization (Windows):**
- *"Create a PivotTable from this data showing total sales by Product, then add a bar chart"*
- *"Import products.csv with Power Query, load to Data Model, create a measure for Total Revenue"*
- *"Create a slicer for the Region field so I can filter the PivotTable interactively"*

**Formatting & Automation (Windows; Mac supports number formats, not the
conditional-formatting or Power Query examples):**
- *"Format the Price column as currency and highlight values over $500 in green"*
- *"Export all Power Query M code to files for version control"*
- *"Show me Excel while you work"* - watch changes in real-time


## Quick Start

1. **Install this extension** (you just did!)
2. **Ask Copilot** in the chat panel:
   - "Open my workbook at this absolute path and read Sheet1!A1:C4"
   - "Create a new workbook at this absolute path and put my data in plain cells"
   - "Add a SUM formula and calculate the workbook"

Supply a full Windows or Mac path and use the returned session ID. On Mac, use
only enabled actions in the [beta support reference](https://github.com/sbroenne/mcp-server-excel/blob/main/specs/MACOS-SUPPORT.md).

**That's it!** The extension includes a self-contained MCP server - no .NET runtime or SDK needed.

| Ask Copilot | Result |
|---|---|
| "Refresh every Power Query query in my sales workbook, then save it." | An updated workbook with fresh query results. [Refresh guide](https://excelmcpserver.dev/guides/refresh-power-query/) |
| "Create a PivotTable showing revenue by region, add a column chart, and save the workbook." | A summary and chart built in Excel. [PivotTable guide](https://excelmcpserver.dev/guides/automate-pivottables/) |
| "Export the Power Query M code and VBA modules from my workbook to source files." | Query and macro source files you can review and keep in version control. [Power Query reference](https://excelmcpserver.dev/reference/powerquery/) |

## Key Features

Excel MCP Server (excel-mcp) provides **60 MCP tools across 31 feature areas, with 388 operations**:

- **Power Query & M code** - Import data, create and edit queries, refresh results, and choose load destinations.
- **Power Pivot & DAX** - Build Data Models, create measures, and manage relationships.
- **PivotTables, charts & slicers** - Build and format reports from tables, ranges, or the Data Model.
- **Data & workbooks** - Read and write data, formulas, and formatting; filter and sort tables; manage worksheets, named ranges, and connections.
- **VBA macros** - Inspect, import, update, run, and export VBA modules.
- **Python in Excel** - Write and run `=PY()` formulas using Excel's cloud Python engine.
- **Watch Copilot work** - See Excel side by side with live status feedback.

[See all 60 MCP tools and 388 operations](https://excelmcpserver.dev/features/).
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

- **Windows with Microsoft Excel 2016+**, or **Apple Silicon macOS with Excel for Mac 16.112+ (experimental beta)**
- **Interactive desktop Excel** must be installed; Intel Macs, Linux, and headless hosts are unsupported

## Potential Issues

**"Excel is not installed" error:**
- Ensure the required desktop Excel version for your platform is installed
- Try opening Excel manually to verify it works

**"VBA access denied" error (Windows only):**
- VBA project inspection/editing requires one-time manual setup in Excel;
  running an existing macro has separate macro-security requirements
- Go to: File → Options → Trust Center → Trust Center Settings → Macro Settings
- Check "Trust access to the VBA project object model"

VBA is unsupported in the Mac beta; changing macro trust cannot enable it.

**Mac platform or permission errors:**
- `PlatformNotSupported`: use an enabled action or the Windows backend
- Grant Excel Automation permission manually when macOS requests it
- `RecoveryRequired`: reconcile the exact workbook manually before retrying;
  do not close unrelated workbooks or terminate shared Excel

**Copilot doesn't see Excel tools:**
- Restart VS Code after installing the extension

### Troubleshooting

- Check Output panel → "Excel MCP Server" for connection status

## Documentation & Support

- **[Complete Documentation](https://excelmcpserver.dev/)** - Full guides and examples
- **[Report Issues](https://github.com/sbroenne/mcp-server-excel/issues)** - Bug reports and feature requests

## License & Privacy

MIT License - see [LICENSE](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)

Privacy Policy - see [PRIVACY.md](https://github.com/sbroenne/mcp-server-excel/blob/main/PRIVACY.md)

---

**Built with GitHub Copilot** | **Powered by Model Context Protocol**
