# Excel MCP Server - AI-Powered Excel Automation

[![GitHub](https://img.shields.io/badge/GitHub-sbroenne%2Fmcp--server--excel-blue)](https://github.com/sbroenne/mcp-server-excel)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Control Microsoft Excel with AI through GitHub Copilot - just ask in natural language!**

Instead of clicking through Excel menus, simply ask:

**Data Transformation & Analysis:**
- *"Optimize all my Power Queries in this workbook for better performance"*
- *"Create a PivotTable from SalesData table showing top 10 products by region"*
- *"Build a DAX measure calculating year-over-year growth with proper time intelligence"*

**Formatting & Styling (No Programming Required):**
- *"Format revenue columns as currency, make headers bold with blue background"*
- *"Apply conditional formatting to highlight values above $10,000 in red"*
- *"Convert this range to an Excel Table with filters and totals row"*

**Workflow Automation:**
- *"Find all cells containing 'Q1 2024' and replace with 'Q1 2025'"*
- *"Add data validation dropdowns to Status column with options: Active, Pending, Completed"*

**🛡️ 100% Safe - Uses Excel's Native API**

Unlike third-party libraries that manipulate `.xlsx` files directly (risking file corruption), Excel MCP Server uses **Excel's official COM API**. This ensures:
- ✅ **Zero risk of document corruption** - Excel handles all file operations safely
- ✅ **Interactive development** - See changes in real-time as you work with live Excel files
- ✅ **Comprehensive automation** - Currently supports 163 operations across 12 specialized tools

**Quick Example:**

```
You: "Create a Power Query named 'SalesData' that loads from data.csv into a worksheet and data model and creates matching pivot tables. Format it professionally."

Copilot uses Excel MCP Server to:
1. Create/open an Excel workbook
2. Add the Power Query with proper M code based on its analysis of the CSV file
3. Load the data to a worksheet and a data model
4. Create profesionally formatted pivot tables

Result: Just what you requested - without touching Excel once.
```

## 👥 Who Should Use This?

**Perfect for:**
- ✅ **Data analysts** automating repetitive Excel workflows
- ✅ **Developers** building Excel-based data solutions
- ✅ **Business users** managing complex Excel workbooks
- ✅ **Teams** maintaining Power Query/VBA/DAX code in Git

**Not suitable for:**
- ❌ Linux/macOS users (Windows + Excel installation required)
- ❌ High-volume batch operations (consider Excel-free alternatives)

## Quick Start

1. **Install this extension** (you just did!)
2. **Ask Copilot** in the chat panel:
   - "List all Power Query queries in workbook.xlsx"
   - "Create a DAX measure for year-over-year revenue growth"
   - "Export all Powere Queires and VBA modules to .vba files for version control"

**That's it!** The extension automatically installs .NET 8 runtime and includes a bundled MCP server.

➡️ **[Learn more and see examples](https://sbroenne.github.io/mcp-server-excel/)**

## Requirements

- **Windows OS** - Excel COM automation requires Windows
- **Microsoft Excel 2016+** - Must be installed on your system
- **.NET 8 Runtime** - **Automatically installed** by the extension

## What's Included

## Features

The Excel MCP Server provides **12 specialized tools** for comprehensive Excel automation:

| Tool | Operations | Purpose |
|------|------------|---------|
| **excel_powerquery** | 11 actions | Power Query M code: create, view, update, delete, load configuration, list sources |
| **excel_datamodel** | 14 actions | Power Pivot (Data Model): DAX measures, relationships, discover structure (tables, columns), refresh |
| **excel_table** | 26 actions | Excel Tables: lifecycle, columns, filters, sorts, structured references, number formatting |
| **excel_pivottable** | 20 actions | PivotTables: create, field management, aggregations, filters, sorting, extract data |
| **excel_range** | 43 actions | Ranges: get/set values/formulas, formatting, validation, clear, copy, insert/delete, find/replace, merge, cell protection |
| **excel_conditionalformat** | 2 actions | Conditional Formatting: add rules (cell value, expression-based), clear rules |
| **excel_vba** | 7 actions | VBA: list, view, export, import, update, run, delete modules |
| **excel_connection** | 9 actions | Connections: OLEDB/ODBC/Text/Web management, properties, refresh, test |
| **excel_worksheet** | 13 actions | Worksheets: lifecycle, tab colors, visibility (list, create, rename, copy, delete, show/hide, very-hide) |
| **excel_namedrange** | 7 actions | Named ranges: list, get, set, create, delete, update, bulk create |
| **excel_file** | 6 actions | File operations: create empty, open, save, close, close-workbook, test |
| **Total** | **155 actions**

## Troubleshooting

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
