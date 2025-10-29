# ExcelMcp - AI-Powered Excel Automation

[![VS Code Marketplace](https://img.shields.io/visual-studio-marketplace/v/sbroenne.excelmcp?label=VS%20Code%20Marketplace)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excelmcp)
[![Installs](https://img.shields.io/visual-studio-marketplace/i/sbroenne.excelmcp)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excelmcp)
[![GitHub](https://img.shields.io/badge/GitHub-sbroenne%2Fmcp--server--excel-blue)](https://github.com/sbroenne/mcp-server-excel)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Control Microsoft Excel with AI assistants through natural language**

Ask GitHub Copilot to manage Power Query M code, DAX measures, VBA macros, Excel Tables, ranges, worksheets, and data connections—all without leaving VS Code.

## What You Can Do

### Ask Copilot in Chat:

**Power Query Development:**
```
"List all Power Query queries in Sales-Analysis.xlsx"
"Show me the M code for the 'TransformSales' query"
"Refactor this slow query to improve performance"
"Export all queries to .pq files for version control"
```

**Data Model & DAX:**
```
"Create a DAX measure for total revenue with currency format"
"List all measures in the Sales table"
"Show me the formula for 'YoY Growth' measure"
"Create a relationship between Sales[CustomerID] and Customers[ID]"
```

**Excel Tables:**
```
"Create an Excel table named 'SalesData' from range A1:D100"
"Add a filter to the Amount column showing values > 1000"
"Sort the table by Region ascending, then Amount descending"
"Add a new calculated column 'Profit' to the table"
```

**VBA Automation:**
```
"List all VBA modules in Report.xlsm"
"Export the 'DataProcessor' module to a .vba file"
"Add error handling to the 'GenerateReport' macro"
```

**Range Operations:**
```
"Read values from Sheet1 range A1:D10"
"Set formulas in column D to calculate totals"
"Copy formatting from A1:D1 to A100:D100"
"Find all cells containing 'ERROR' in the worksheet"
```

## Quick Start

1. **Install this extension** (you just did!)
2. **Open a workspace** with Excel files or create a new one
3. **Ask Copilot** in the chat panel:
   - `@workspace List all Power Query queries in workbook.xlsx`
   - `@workspace Create a new Excel file with sample data`
   - `@workspace Export all DAX measures to version control`

**That's it!** The extension automatically installs .NET 8 runtime and the MCP server. No manual configuration needed.

## Common Use Cases

### Power Query Refactoring
**Scenario:** You have slow Power Query transformations slowing down your workbook refresh.

**Ask Copilot:** "Analyze the 'TransformCustomerData' query and suggest performance optimizations"

**What happens:** Copilot views the M code, identifies bottlenecks (like late filtering), refactors the query, and updates it in your workbook.

---

### DAX Measure Development
**Scenario:** You need to create complex DAX measures for financial reporting.

**Ask Copilot:** "Create a DAX measure for year-over-year revenue growth with percentage formatting"

**What happens:** Copilot creates the measure with proper CALCULATE syntax, applies percentage format, and adds it to your Data Model.

---

### VBA Version Control
**Scenario:** You want to track VBA macro changes in git.

**Ask Copilot:** "Export all VBA modules to .vba files in the 'vba-modules' folder"

**What happens:** Copilot lists all modules, exports each to individual files, enabling git tracking and code review.

---

### Excel Table Automation
**Scenario:** You need to filter and sort large datasets for analysis.

**Ask Copilot:** "In the SalesTable, filter Amount > 10000 and sort by Date descending"

**What happens:** Copilot applies the filter criteria and multi-level sort, preparing the data for your analysis.

## Requirements

- **Windows OS** - Excel COM automation requires Windows
- **Microsoft Excel 2016+** - Must be installed on your system
- **.NET 8 Runtime** - **Automatically installed** by the extension

## What's Included

The ExcelMcp MCP server provides **10 specialized tools** with 100+ operations:

| Tool | Operations | Purpose |
|------|------------|---------|
| **excel_powerquery** | 11 actions | Create, view, update, refactor M code |
| **excel_datamodel** | 20 actions | DAX measures, relationships, calculated columns |
| **table** | 22 actions | Table lifecycle, columns, filters, sorts |
| **excel_range** | 30+ actions | Read/write values, formulas, formatting |
| **excel_vba** | 7 actions | List, export, import, run VBA code |
| **excel_connection** | 11 actions | Manage OLEDB, ODBC, Text, Web connections |
| **excel_worksheet** | 5 actions | Create, rename, copy, delete sheets |
| **excel_parameter** | 6 actions | Named range management |
| **excel_file** | 1 action | Create Excel workbooks |
| **excel_version** | 1 action | Check for updates |

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

## How It Works

This extension uses the Model Context Protocol (MCP) to connect AI assistants to Excel:

1. The extension registers the ExcelMcp MCP server with VS Code
2. When you ask Copilot about Excel, VS Code runs: `dotnet tool run mcp-excel`
3. The MCP server uses Excel COM automation to perform operations
4. Results are returned to Copilot in your chat

The extension automatically handles .NET installation via the .NET Install Tool.

## Documentation & Support

- **[Complete Documentation](https://github.com/sbroenne/mcp-server-excel)** - Full guides and API reference
- **[MCP Server Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.McpServer/README.md)** - Detailed tool reference with examples
- **[Command Reference](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/COMMANDS.md)** - All 100+ operations documented
- **[Report Issues](https://github.com/sbroenne/mcp-server-excel/issues)** - Bug reports and feature requests
- **[Discussions](https://github.com/sbroenne/mcp-server-excel/discussions)** - Community support

## License

MIT License - see [LICENSE](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)

---

**Built with GitHub Copilot** | **Powered by Model Context Protocol** | **Excel COM Automation**
