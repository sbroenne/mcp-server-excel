# ExcelMcp - MCP Server for Microsoft Excel

[![Build MCP Server](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml)
[![Build CLI](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total)](https://github.com/sbroenne/mcp-server-excel/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-8.0-blue.svg)](https://dotnet.microsoft.com/download/dotnet/8.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

**A Model Context Protocol (MCP) server that gives AI assistants full control over Microsoft Excel through native COM automation.**

Control Power Query M code, Power Pivot (Data Model with DAX measures and relationships), VBA macros, Excel Tables, PivotTables, connections, ranges, and worksheets through conversational AI. Also includes a CLI for direct human automation and for tools that do not support using a MCP Server.

## 🤔 What is This?

**ExcelMcp lets you control Excel using conversational AI (like GitHub Copilot or Claude).**

Instead of clicking through Excel menus or writing complex VBA, you can simply ask:
- *"Create a Power Query that combines Sales.csv and Products.csv on ProductID"*
- *"Add a DAX measure in Power Pivot calculating year-over-year revenue growth"*
- *"Create a PivotTable showing sales by region and product category"*
- *"Export all VBA modules to separate files for Git version control"*
- *"Create a table with filters and sort by Revenue descending"*

The AI assistant uses this MCP server to execute your requests **directly in your Excel application** - no manual clicking required.

**Quick Example:**

```
You: "Create a Power Query named 'SalesData' that loads from data.csv"

AI Assistant uses ExcelMcp to:
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
- ❌ Server-side data processing (use libraries like ClosedXML, EPPlus instead)
- ❌ Linux/macOS users (Windows + Excel installation required)
- ❌ High-volume batch operations (consider Excel-free alternatives)

## 🎯 What You Can Do

**Development & Automation:**
- 🔄 **Power Query** - 12 operations: create/edit M code, manage transformations, load configurations (worksheet, data model, connection only)
- 📊 **Power Pivot (Data Model)** - 13 operations: build DAX measures, manage relationships, export to .dax files
- 🎨 **Excel Tables** - 24 operations: automate formatting, filtering, sorting, structured references, number formats
- 📈 **PivotTables** - 17 operations: create and configure PivotTables for interactive analysis
- 📝 **VBA Macros** - 7 operations: export/import/run VBA code, integrate with version control
- 📋 **Ranges & Data** - 31 operations: values, formulas, copy/paste, find/replace, formatting, validation
- 📄 **Worksheets** - 12 operations: lifecycle management, tab colors, visibility controls
- 🔌 **Connections** - 11 operations: manage OLEDB, ODBC, Text, Web data sources
- 🏷️ **Parameters** - 7 operations: named range management and bulk operations

**AI-Powered Workflows:**
- 💬 Talk to Excel in natural language through GitHub Copilot or Claude
- 🤖 Automate repetitive Excel tasks with conversational commands
- 📦 Version control Excel code artifacts (Power Query, VBA, DAX measures)
- 🔄 Build data pipelines with AI assistance

<details>
<summary>📚 <strong>See Complete Feature List (100+ Operations)</strong></summary>

### Power Query & M Code (12 operations)
- Create, read, update, delete Power Query transformations
- Export/import M code for version control
- Manage query load destinations (worksheet/data model/connection-only/both)
- Set privacy levels for data source combinations
- Get load configuration for existing queries

### Data Model & DAX (Power Pivot) (13 operations)
- Create/update/delete DAX measures with format types (Currency, Percentage, Decimal, General)
- Manage table relationships (create, toggle active/inactive, delete)
- Discover model structure (tables, columns, measures, relationships)
- Export measures to .dax files for Git workflows
- Get comprehensive model information
- **Note:** DAX calculated columns are not supported (use Excel UI for calculated columns)

### Excel Tables (ListObjects) (24 operations)
- Lifecycle: create, resize, rename, delete, get info
- Styling: apply table styles, toggle totals row, set column totals
- Column management: add, remove, rename columns
- Data operations: append rows, apply filters (criteria/values), clear filters, get filter state
- Sorting: single-column sort, multi-column sort (up to 3 levels)
- Number formatting: get/set column number formats
- Advanced features: structured references, Data Model integration

### PivotTables (17 operations)
- Creation: create from ranges or Excel Tables
- Field management: add/remove fields to Row, Column, Value, Filter areas
- Aggregation functions: Sum, Average, Count, Min, Max, etc. with validation
- Advanced features: field filters, sorting, custom field names, number formatting
- Data extraction: get PivotTable data as 2D arrays for further analysis
- Lifecycle: list, get info, delete, refresh

### VBA Macros (7 operations)
- List all VBA modules and procedures
- View module code without exporting
- Export/import VBA modules to/from files
- Update existing modules
- Execute macros with parameters
- Delete modules
- Version control VBA code through file exports

### Ranges & Worksheets
- **Data Operations** (10 actions): get/set values/formulas, clear (all/contents/formats), copy/paste (all/values/formulas), insert/delete rows/columns/cells, find/replace, sort
- **Number Formatting** (3 actions): get formats as 2D arrays, apply uniform format, set individual cell formats
- **Visual Formatting** (1 action): font (name, size, bold, italic, underline, color), fill color, borders (style, weight, color), alignment (horizontal, vertical, wrap text, orientation)
- **Data Validation** (1 action): dropdown lists, number/date/text rules, operators (Between, Greater, Equal, etc.), input messages, error alerts
- **Hyperlinks** (4 actions): add, remove, list all, get specific hyperlink
- **Smart Range Operations** (3 actions): UsedRange, CurrentRegion, get range info (address, dimensions, format)
- **31 range operations total** covering all common Excel range manipulation needs
- **Worksheet management** (12 actions): lifecycle (create, rename, copy, delete), tab colors (set, get, clear), visibility controls (show, hide, very-hide, get status)

### Data Connections (11 operations)
- Manage OLEDB, ODBC, Text, Web connections
- Import/export connections via .odc files
- Update connection strings and properties
- Refresh connections and test connectivity
- Load connection-only connections to worksheet tables
- Get/set connection properties (refresh settings, background query, etc.)

### Named Ranges/Parameters (7 operations)
- List all named ranges with references
- Get/set single values
- Create/delete named ranges
- Update cell references
- Bulk create multiple parameters

</details>

## 🚀 Quick Start (2 Minutes)

**Requirements:** Windows OS + Microsoft Excel 2016+

### ⭐ Recommended: VS Code Extension (One-Click Setup)

**Fastest way to get started - everything configured automatically:**

1. **Install Extension**
   - Open VS Code → Extensions (`Ctrl+Shift+X`)
   - Search for **"ExcelMcp"**
   - Click **Install**

2. **Automatic Setup** (no manual steps!)
   - ✅ Installs .NET 8 runtime
   - ✅ Includes bundled MCP server
   - ✅ Registers with AI assistants
   - ✅ Shows quick start guide

3. **Start Using It**
   
   The extension opens automatically after installation with a quick start guide!

---

### Manual Installation (Advanced Users)

**For non-VS Code environments or manual setup:**

```powershell
# Install .NET 8 SDK
winget install Microsoft.DotNet.SDK.8

# Install ExcelMcp MCP server as a global tool
dotnet tool install --global Sbroenne.ExcelMcp.McpServer

# To update to the latest version later:
dotnet tool update --global Sbroenne.ExcelMcp.McpServer
```

**Configure Your AI Assistant**

**For GitHub Copilot in VS Code** - Create `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "mcp-excel"]
    }
  }
}
```

**For GitHub Copilot in Visual Studio** - Create `.mcp.json` in your solution directory or `%USERPROFILE%\.mcp.json`:

```json
{
  "servers": {
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "mcp-excel"]
    }
  }
}
```

**For Claude Desktop** - Add to your MCP configuration:

```json
{
  "mcpServers": {
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "mcp-excel"]
    }
  }
}
```

**Test It Out**

Try a practical example - ask your AI assistant:
```
Create an empty Excel file called "test.xlsx" and add a Power Query that loads data from a CSV file
```

The AI will guide you through the process and execute the commands directly!

---

## 🔧 How It Works - COM Interop Architecture

**ExcelMcp uses Windows COM automation to control the actual Excel application (not just .xlsx files).**

This means you get:
- ✅ **Full Excel Feature Access** - Power Query engine, VBA runtime, Data Model, calculation engine, pivot tables
- ✅ **True Compatibility** - Works exactly like Excel UI, no feature limitations
- ✅ **Live Data Operations** - Refresh Power Query, connections, Data Model in real workbooks
- ✅ **Interactive Development** - Immediate Excel feedback as AI makes changes
- ✅ **All File Formats** - .xlsx, .xlsm, .xlsb, even legacy formats

**Technical Requirements:**
- ⚠️ **Windows Only** - COM interop is Windows-specific
- ⚠️ **Excel Required** - Microsoft Excel 2016 or later must be installed
- ⚠️ **Desktop Environment** - Controls actual Excel process (not for server-side processing)

## 🔟 MCP Tools Overview

**11 specialized tools for comprehensive Excel automation:**

1. **excel_powerquery** (11 actions) - Power Query M code: create, view, import, export, update, delete, manage load destinations, privacy levels
2. **excel_datamodel** (14 actions) - Power Pivot (Data Model): CRUD DAX measures/relationships, discover structure, export to .dax files
3. **excel_table** (22 actions) - Excel Tables: lifecycle, columns, filters, sorts, structured references, totals, Data Model integration
4. **excel_pivottable** (20 actions) - PivotTables: create from ranges/tables, field management (row/column/value/filter), aggregations, filters, sorting, extract data
5. **excel_range** (38+ actions) - Ranges: get/set values/formulas, number formatting, visual formatting (font, fill, border, alignment), data validation, clear, copy, insert/delete, find/replace, sort, hyperlinks
6. **excel_vba** (7 actions) - VBA: list, view, export, import, update, run, delete modules
7. **excel_connection** (11 actions) - Connections: OLEDB/ODBC/Text/Web management, properties, refresh, test
8. **excel_worksheet** (13 actions) - Worksheets: lifecycle (list, create, rename, copy, delete), tab colors (set-tab-color, get-tab-color, clear-tab-color), visibility (set-visibility, get-visibility, show, hide, very-hide)
9. **excel_parameter** (6 actions) - Named ranges: list, get, set, create, delete, update
10. **excel_file** (1 action) - File creation: create empty .xlsx/.xlsm workbooks
11. **Batch Session Tools** (3 actions) - Multi-operation performance: begin-batch, execute-in-batch, commit-batch

> 📚 **[Complete MCP Server Guide →](src/ExcelMcp.McpServer/README.md)** - Detailed tool documentation and examples

---

## 📋 Additional Information

### CLI for Direct Automation

ExcelMcp also provides a command-line interface for vba-based Excel automation (no AI required). See **[CLI Guide](docs/CLI.md)** for complete documentation.

### Project Information

**License:** MIT License - see [LICENSE](LICENSE) file

**Contributing:** See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for guidelines

**Built With:** This entire project was developed using GitHub Copilot AI assistance

**Acknowledgments:**
- Microsoft Excel Team - For comprehensive COM automation APIs
- Model Context Protocol community - For the AI integration standard
- Open Source Community - For inspiration and best practices

---

### SEO & Discovery

`MCP Server` • `Model Context Protocol` • `Excel Automation` • `GitHub Copilot` • `AI Excel` • `Power Query` • `Power Pivot` • `DAX Measures` • `Data Model` • `PivotTables` • `VBA Macros` • `Excel Tables` • `COM Interop` • `Windows Excel` • `Excel Development`
