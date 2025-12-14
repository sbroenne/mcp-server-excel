# Excel MCP Server - AI-Powered Excel Automation

[![GitHub](https://img.shields.io/badge/GitHub-sbroenne%2Fmcp--server--excel-blue)](https://github.com/sbroenne/mcp-server-excel)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)


**Control Microsoft Excel with AI through GitHub Copilot - just ask in natural language!**

**MCP Server for Excel** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands. Automate Power Query, DAX measures, VBA macros, PivotTables, Charts, formatting, and data transformations - no Excel programming knowledge required. 

**🛡️ 100% Safe - Uses Excel's Native COM API** - Zero risk of file corruption. Unlike third-party libraries that manipulate `.xlsx` files directly, this project uses Excel's official API ensuring complete safety and compatibility.

**💡 Interactive Development** - See results instantly in Excel. Create a query, run it, inspect the output, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing.

## Features

The Excel MCP Server provides **12 specialized tools with 180 operations** for comprehensive Excel automation:

- 🔄 **Power Query** (9 ops) - Atomic workflows, M code management, load destinations
- 📊 **Data Model/DAX** (16 ops) - Measures, relationships, model structure
- 🎨 **Excel Tables** (24 ops) - Lifecycle, filtering, sorting, structured references
- 📈 **PivotTables** (30 ops) - Creation, fields, aggregations, calculated members/fields
- 📉 **Charts** (14 ops) - Create, configure, manage series and formatting
- 📝 **VBA** (6 ops) - Modules, execution, version control
- 📋 **Ranges** (42 ops) - Values, formulas, formatting, validation, protection
- 📄 **Worksheets** (16 ops) - Lifecycle, colors, visibility, cross-workbook moves
- 🔌 **Connections** (9 ops) - OLEDB/ODBC management and refresh
- 🏷️ **Named Ranges** (6 ops) - Parameters and configuration
- 📁 **Files** (6 ops) - Session management and workbook creation
- 🎨 **Conditional Formatting** (2 ops) - Rules and clearing

📚 **[Complete Feature Reference →](https://github.com/sbroenne/mcp-server-excel/blob/main/FEATURES.md)**


## 💬 Example Prompts

**Data Transformation & Analysis:**
- *"Optimize all my Power Queries in this workbook for better performance"*
- *"Create a PivotTable from SalesData table showing top 10 products by region with sum and average"*
- *"Create a data model from the following tables ... "*
- *"Build a DAX measure calculating year-over-year growth with proper time intelligence"*
- *"Filter this table by Column Product = Sushi"*
- *"Create a treemap chart from this table".

**Formatting & Styling (No Programming Required):**
- *"Format the revenue columns as currency, make headers bold with blue background, and add borders to the table"*
- *"Apply conditional formatting to highlight values above $10,000 in red and below $5,000 in yellow"*
- *"Convert this data range to an Excel Table with style TableStyleMedium2, add auto-filters, and create a totals row"*

**Workflow Automation:**
- *"Find all cells containing 'Q1 2024' and replace with 'Q1 2025', then sort the table by Date descending"*
- *"Add data validation dropdowns to the Status column with options: Active, Pending, Completed"*
- *"Merge the header cells, center-align them, and auto-fit all column widths to content"*


## Quick Start

1. **Install this extension** (you just did!)
2. **Ask Copilot** in the chat panel:
   - "List all Power Query queries in workbook.xlsx"
   - "Create a DAX measure for year-over-year revenue growth"
   - "Export all Powere Queires and VBA modules to .vba files for version control"

**That's it!** The extension automatically installs .NET 8 runtime and includes a bundled MCP server.

➡️ **[Learn more and see examples](https://sbroenne.github.io/mcp-server-excel/)**

## Status Bar Monitor

The extension displays a **status bar item** labeled "Excel MCP" to show server connection status and manage sessions.

**Status Indicators:**
- ✅ `Excel MCP (N)` - Connected with N active sessions
- ⚠️ `Excel MCP (Disconnected)` - Server not running or unreachable

**Session Management:**
- **Click the status bar** to view active sessions in a Quick Pick menu
- **Select a session** to choose an action:
  - **Close** - Terminate session without saving
  - **Save & Close** - Save changes then close
- **Hover** over the status bar to see session count and status summary

**Command Palette:**
- `Excel MCP: Show Sessions` - Open session manager (also via status bar click)

**Configuration:**
- `excelMcp.pollIntervalMs` - Status update interval in milliseconds (default: 3000, range: 1000-60000)

## Requirements

- **Windows OS** - Excel COM automation requires Windows
- **Microsoft Excel 2016+** - Must be installed on your system
- **.NET 8 Runtime** - **Automatically installed** by the extension

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
- ### Troubleshooting

- Check Output panel → "Excel MCP Server" for connection status

## Documentation & Support

- **[Complete Documentation](https://github.com/sbroenne/mcp-server-excel)** - Full guides and examples
- **[Report Issues](https://github.com/sbroenne/mcp-server-excel/issues)** - Bug reports and feature requests

## License

MIT License - see [LICENSE](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)

---

**Built with GitHub Copilot** | **Powered by Model Context Protocol**
