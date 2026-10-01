# ExcelMcp - Real Excel Automation for VS Code

[![GitHub](https://img.shields.io/badge/GitHub-sbroenne%2Fmcp--server--excel-blue)](https://github.com/sbroenne/mcp-server-excel)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)


**Automate real Microsoft Excel from VS Code with GitHub Copilot.**

Ask Copilot to refresh Power Query, calculate formulas, create PivotTables and
charts, work with DAX measures, or run VBA. This extension includes the MCP
server and its Excel guidance skill in one installation.

**⚡ Powered by the real Excel engine** - ExcelMcp automates the **actual Excel application** through its official COM API — the same engine Excel itself uses. That unlocks what spreadsheets are really for:

- **Runs live Excel operations** - Refresh Power Query to pull and reshape fresh data, recalculate with Excel's own engine, refresh PivotTables and the Data Model, evaluate DAX, and run VBA or Python `=PY()` — the real, *computed results* land right in your workbook.
- **Edits your existing files safely** - Excel opens and saves the workbook itself, so every formula, PivotTable, chart, macro, the Data Model and all your formatting stay exactly as they were.

File-parser tools cannot run Excel's calculation, Power Query, or Data Model
engines. ExcelMcp uses Excel itself, so you can inspect results live and keep
editing the workbook normally. Just say *"Show me Excel while you work."*

[![A sales table, regional summary, and chart created in real Microsoft Excel](https://excelmcpserver.dev/assets/images/excel-demo-table-chart.png)](https://excelmcpserver.dev/use-cases/)

**💡 Interactive Development** - See results instantly in Excel. Create a query, run it, inspect the output, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing.

## Key features

The Excel MCP Server (excel-mcp) provides **31 specialized tools with 326 operations** for comprehensive Excel automation:

- 🔄 **Power Query & M code** - Create, edit and optimize M code. Import from files, databases and APIs. Refresh queries and manage load destinations.
- 🧮 **Power Pivot & DAX** - Build Data Models, create DAX measures and manage table relationships. Full Power Pivot automation.
- 📊 **PivotTables & charts** - Create PivotTables from ranges, tables or the Data Model. Build charts and PivotCharts with full formatting control.
- 📋 **Tables & ranges** - Read/write data, formulas and formatting. Filter, sort and validate. Manage Excel Tables with structured references.
- 📝 **VBA macros** - View, import, update and execute VBA code. Export modules for version control.
- 📄 **Worksheets & connections** - Manage sheets, named ranges and data connections. Copy and move sheets between workbooks.
- 👁️ **Agent mode** - Watch AI work in Excel in real time — side-by-side view, live status-bar feedback and smart window arrangement, like a pair programmer in a spreadsheet.
- 🐍 **Python in Excel** - Write and run `=PY()` formulas that execute in Excel's cloud Python engine — process worksheet data with pandas, NumPy and more, from your AI assistant.
- 🧪 **LLM-tested quality** - Tool behavior validated with real LLM workflows using [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering), so AI assistants reliably understand and use every operation.

📚 **[See all 31 tools and 326 operations →](https://excelmcpserver.dev/features/)**

### Agent Skills (Bundled)

This extension includes an **Agent Skill** following the [agentskills.io](https://agentskills.io) specification - providing domain-specific guidance for AI assistants:

- **[excel-mcp](https://excelmcpserver.dev/skills/)** - MCP Server tool guidance

The skill is registered automatically through VS Code's `chatSkills`
contribution point. No separate skill installation or preview setting is needed.
Copilot can load relevant guidance as needed. Type `/skills` in chat to open
VS Code's Configure Skills menu. Ordinary natural-language requests work
without remembering a command.


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

1. **Install the extension** on your Windows desktop with Excel installed.
2. **Open a Copilot chat that can use tools.** Run **MCP: List Servers** from
   the Command Palette, select **excel-mcp**, and start it. Approve the server
   and tool use when VS Code asks.
3. **Ask about a workbook**, using a file available on your Windows machine:
   - "List the Power Query queries in my sales workbook."
   - "Create a PivotTable showing revenue by region, then add a column chart."
   - "Export my Power Query M code and VBA modules for version control."

The extension includes a self-contained MCP server - **no separate .NET,
Node.js, CLI, or skill installation is needed**.

Close a workbook in Excel before asking Copilot to open it: ExcelMcp needs
exclusive access while automating it.

➡️ **[Learn more and see examples](https://excelmcpserver.dev/)**

## Requirements

- **Windows x64 or Windows ARM64** with an interactive desktop. On ARM64,
  the bundled x64 server runs through Windows' x64 emulation.
- **Microsoft Excel 2016 or later**, installed and able to open normally.
  Some features, such as Python in Excel, require a supported Excel edition.
- **VS Code 1.125 or later** and GitHub Copilot chat with tool support.

This extension is not for macOS, Linux, browser-only VS Code, Windows services,
or unattended server-side processing. It bundles the MCP server, not `excelcli`.

### Remote workspaces

Excel runs on your **local Windows desktop**, even when VS Code is connected
to WSL, SSH, a container, or a Codespace. A remote workspace path is not a local
Excel file path. Copy or synchronize the workbook to your Windows machine
before opening it with ExcelMcp. Remote connections do not add Linux or
server-side Excel support.

## Potential Issues

**"Excel is not installed" error:**
- Ensure Microsoft Excel 2016+ is installed on your Windows machine
- Try opening Excel manually to verify it works

**"VBA access denied" error:**
- VBA operations require one-time manual setup in Excel
- Go to: File → Options → Trust Center → Trust Center Settings → Macro Settings
- Check "Trust access to the VBA project object model"

**Copilot doesn't see Excel tools:**
- Run **MCP: List Servers**, choose **excel-mcp**, and start it.
- Check that Excel tools are enabled in your Copilot chat.
- Accept the server trust prompt if you want to use this bundled server.
- After an extension update, refresh the tools when VS Code prompts you.

**"Bundled server is missing or unreadable" error:**
- Check that security software has not blocked the bundled executable.
- Check file permissions or reinstall the extension.

**"Could not check Excel registration" or a registration timeout:**
- Verify Windows PowerShell and desktop Excel open normally.
- Retry; repair Microsoft Office if Excel's installation is damaged.

### Troubleshooting

- **Server logs:** run **MCP: List Servers**, choose **excel-mcp**, then
  **Show Output**.
- **Extension setup diagnostics:** open the Output panel and choose
  **ExcelMcp**, or select **Show Setup Output** in a setup error notification.
- Startup checks read Excel's registration; they do not start Excel, open a
  workbook, or verify that every Excel feature is available.

## Documentation & Support

- **[User guides](https://excelmcpserver.dev/guides/)** - Practical walkthroughs for using ExcelMcp; also opened by Getting Started
- **[Complete Documentation](https://excelmcpserver.dev/)** - Full guides and examples
- **[Report Issues](https://github.com/sbroenne/mcp-server-excel/issues)** - Bug reports and feature requests

## License & Privacy

MIT License - see [LICENSE](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)

Privacy Policy - see [PRIVACY.md](https://github.com/sbroenne/mcp-server-excel/blob/main/PRIVACY.md)

---

**Built with GitHub Copilot** | **Powered by Model Context Protocol**
