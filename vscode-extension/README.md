# Excel MCP Server - AI-Powered Excel Automation

[![GitHub](https://img.shields.io/badge/GitHub-sbroenne%2Fmcp--server--excel-blue)](https://github.com/sbroenne/mcp-server-excel)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)


**Control Microsoft Excel with AI through GitHub Copilot - just ask in natural language!**

**MCP Server for Excel** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands. Automate Power Query, DAX measures, VBA macros, PivotTables, Charts, formatting, and data transformations - no Excel programming knowledge required. 

**⚡ Powered by the real Excel engine** - ExcelMcp automates the **actual Excel application** through its official COM API — the same engine Excel itself uses. That unlocks what spreadsheets are really for:

- **Runs live Excel operations** - Refresh Power Query to pull and reshape fresh data, recalculate with Excel's own engine, refresh PivotTables and the Data Model, evaluate DAX, and run VBA or Python `=PY()` — the real, *computed results* land right in your workbook.
- **Edits your existing files safely** - Excel opens and saves the workbook itself, so every formula, PivotTable, chart, macro, the Data Model and all your formatting stay exactly as they were.

Other tools (openpyxl-based MCP servers and Agent Skills, including Anthropic's `xlsx` skill) read and rewrite the `.xlsx` file directly — which can quietly drop PivotTables, charts, and macros, and can't run Power Query, the Data Model, or DAX at all. Here, Excel does the work. Watch it live: just say *"Show me Excel while you work."*

**💡 Interactive Development** - See results instantly in Excel. Create a query, run it, inspect the output, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing.

## Key features

The Excel MCP Server (excel-mcp) provides **26 specialized tools with 232 operations** for comprehensive Excel automation:

- 🔄 **Power Query & M code** - Create, edit and optimize M code. Import from files, databases and APIs. Refresh queries and manage load destinations.
- 🧮 **Power Pivot & DAX** - Build Data Models, create DAX measures and manage table relationships. Full Power Pivot automation.
- 📊 **PivotTables & charts** - Create PivotTables from ranges, tables or the Data Model. Build charts and PivotCharts with full formatting control.
- 📋 **Tables & ranges** - Read/write data, formulas and formatting. Filter, sort and validate. Manage Excel Tables with structured references.
- 📝 **VBA macros** - View, import, update and execute VBA code. Export modules for version control.
- 📄 **Worksheets & connections** - Manage sheets, named ranges and data connections. Copy and move sheets between workbooks.
- 👁️ **Agent mode** - Watch AI work in Excel in real time — side-by-side view, live status-bar feedback and smart window arrangement, like a pair programmer in a spreadsheet.
- 🐍 **Python in Excel** - Write and run `=PY()` formulas that execute in Excel's cloud Python engine — process worksheet data with pandas, NumPy and more, from your AI assistant.
- 🧪 **LLM-tested quality** - Tool behavior validated with real LLM workflows using [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering), so AI assistants reliably understand and use every operation.

📚 **[See all 26 tools and 232 operations →](https://excelmcpserver.dev/features/)**

### Agent Skills (Bundled)

This extension includes an **Agent Skill** following the [agentskills.io](https://agentskills.io) specification - providing domain-specific guidance for AI assistants:

- **[excel-mcp](https://excelmcpserver.dev/skills/)** - MCP Server tool guidance

**VS Code setup:** Enable the preview setting `chat.useAgentSkills` to allow Copilot to load skills. Skills are registered via VS Code's `chatSkills` contribution point and managed automatically.


## 💬 Example Prompts

**Create & Populate Data:**
- *"Create a new Excel file called SalesTracker.xlsx with a table for Date, Product, Quantity, Unit Price, and Total"*
- *"Put this data in A1:C4 - Name, Age, City / Alice, 30, Seattle / Bob, 25, Portland"*
- *"Add sample data and a formula column for Quantity times Unit Price"*

**Analysis & Visualization:**
- *"Create a PivotTable from this data showing total sales by Product, then add a bar chart"*
- *"Import products.csv with Power Query, load to Data Model, create a measure for Total Revenue"*
- *"Create a slicer for the Region field so I can filter the PivotTable interactively"*

**Formatting & Automation:**
- *"Format the Price column as currency and highlight values over $500 in green"*
- *"Export all Power Query M code to files for version control"*
- *"Show me Excel while you work"* - watch changes in real-time


## Quick Start

1. **Install this extension** (you just did!)
2. **Ask Copilot** in the chat panel:
   - "List all Power Query queries in workbook.xlsx"
   - "Create a DAX measure for year-over-year revenue growth"
   - "Export all Power Queries and VBA modules to .vba files for version control"

**That's it!** The extension includes a self-contained MCP server - no .NET runtime or SDK needed.

➡️ **[Learn more and see examples](https://excelmcpserver.dev/)**

## Requirements

- **Windows OS** - Excel COM automation requires Windows
- **Microsoft Excel 2016+** - Must be installed on your system

## Potential Issues

**"Excel is not installed" error:**
- Ensure Microsoft Excel 2016+ is installed on your Windows machine
- Try opening Excel manually to verify it works

**"VBA access denied" error:**
- VBA operations require one-time manual setup in Excel
- Go to: File → Options → Trust Center → Trust Center Settings → Macro Settings
- Check "Trust access to the VBA project object model"

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
