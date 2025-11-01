# ExcelMcp - AI-Powered Excel Automation

[![VS Code Marketplace](https://img.shields.io/visual-studio-marketplace/v/sbroenne.excelmcp?label=VS%20Code%20Marketplace)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excelmcp)
[![Installs](https://img.shields.io/visual-studio-marketplace/i/sbroenne.excelmcp)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excelmcp)
[![GitHub](https://img.shields.io/badge/GitHub-sbroenne%2Fmcp--server--excel-blue)](https://github.com/sbroenne/mcp-server-excel)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Control Microsoft Excel with AI through GitHub Copilot - just ask in natural language!**

Instead of clicking through Excel menus or writing complex VBA, simply ask:
- *"Create a Power Query that combines Sales.csv and Products.csv on ProductID"*
- *"Add a DAX measure in Power Pivot calculating year-over-year revenue growth"*
- *"Create a PivotTable showing sales by region and product category"*
- *"Export all VBA modules to separate files for Git version control"*
- *"Create a table with filters and sort by Revenue descending"*

**Quick Example:**

```
You: "Create a Power Query named 'SalesData' that loads from data.csv"

Copilot uses ExcelMcp to:
1. Create/open an Excel workbook
2. Add the Power Query with proper M code
3. Load the data to a worksheet
4. Save and return confirmation

Result: A working Excel file with the query ready to use
```

**🛡️ 100% Safe - Uses Excel's Native API**

Unlike third-party libraries that manipulate `.xlsx` files directly (risking file corruption), ExcelMcp uses **Excel's official COM API**. This ensures:
- ✅ **Zero risk of document corruption** - Excel handles all file operations safely
- ✅ **Interactive development** - See changes in real-time as you work with live Excel files
- ✅ **Growing feature set** - Currently supports 80+ operations across Power Query, Power Pivot, VBA, PivotTables, Tables, and more (active development)

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
   - "Export all VBA modules to .vba files for version control"

**That's it!** The extension automatically installs .NET 8 runtime and includes a bundled MCP server.

## Requirements

- **Windows OS** - Excel COM automation requires Windows
- **Microsoft Excel 2016+** - Must be installed on your system
- **.NET 8 Runtime** - **Automatically installed** by the extension

## What's Included

The ExcelMcp MCP server provides **11 specialized tools** for comprehensive Excel automation:

| Tool | Operations | Purpose |
|------|------------|---------|
| **excel_powerquery** | 11 actions | Power Query M code: create, view, import, export, update, delete |
| **excel_datamodel** | 14 actions | Power Pivot (Data Model): DAX measures, relationships, discover structure |
| **excel_table** | 22 actions | Excel Tables: lifecycle, columns, filters, sorts, structured references |
| **excel_pivottable** | 20 actions | PivotTables: create, field management, aggregations, filters, sorting |
| **excel_range** | 30+ actions | Ranges: get/set values/formulas, clear, copy, insert/delete, find/replace |
| **excel_vba** | 7 actions | VBA: list, view, export, import, update, run, delete modules |
| **excel_connection** | 11 actions | Connections: OLEDB/ODBC/Text/Web management, properties, refresh |
| **excel_worksheet** | 5 actions | Worksheets: list, create, rename, copy, delete |
| **excel_parameter** | 6 actions | Named ranges: list, get, set, create, delete, update |
| **excel_file** | 1 action | File creation: create empty .xlsx/.xlsm workbooks |
| **Batch Session Tools** | 3 actions | Multi-operation performance: begin-batch, execute-in-batch, commit-batch |

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
- Check Output panel → "ExcelMcp" for connection status

## Documentation & Support

- **[Complete Documentation](https://github.com/sbroenne/mcp-server-excel)** - Full guides and examples
- **[Report Issues](https://github.com/sbroenne/mcp-server-excel/issues)** - Bug reports and feature requests

## License

MIT License - see [LICENSE](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)

---

**Built with GitHub Copilot** | **Powered by Model Context Protocol**
