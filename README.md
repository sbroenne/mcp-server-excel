# ExcelMcp - MCP Server for Microsoft Excel

[![Build MCP Server](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml)
[![Build CLI](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total)](https://github.com/sbroenne/mcp-server-excel/releases)

[![NuGet MCP Server](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg?label=MCP%20Server)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![NuGet CLI](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.CLI.svg?label=CLI)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)
[![NuGet Core](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.Core.svg?label=Core)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.Core)
[![NuGet ComInterop](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.ComInterop.svg?label=ComInterop)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.ComInterop)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-8.0-blue.svg)](https://dotnet.microsoft.com/download/dotnet/8.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

**Automate Excel with AI - A Model Context Protocol (MCP) server for comprehensive Excel automation through conversational AI.**

Talk to Excel using GitHub Copilot or Claude to automate Power Query M code, Power Pivot DAX measures, PivotTables, Excel Tables, formatting, data transformations, and complex analysis - all through natural language commands. No programming required. Built on Excel's native COM API for zero risk of file corruption.

## 🤔 What is This?

**ExcelMcp connects AI assistants (GitHub Copilot, Claude) to Microsoft Excel for hands-free automation of complex Excel tasks.**

Stop manually clicking through Excel menus for repetitive tasks. Instead, describe what you want in plain English:

**Data Transformation & Analysis:**
- *"Optimize all my Power Queries in this workbook for better performance"*
- *"Create a PivotTable from SalesData table showing top 10 products by region with sum and average"*
- *"Build a DAX measure calculating year-over-year growth with proper time intelligence"*

**Formatting & Styling (No Programming Required):**
- *"Format the revenue columns as currency, make headers bold with blue background, and add borders to the table"*
- *"Apply conditional formatting to highlight values above $10,000 in red and below $5,000 in yellow"*
- *"Convert this data range to an Excel Table with style TableStyleMedium2, add auto-filters, and create a totals row"*

**Workflow Automation:**
- *"Find all cells containing 'Q1 2024' and replace with 'Q1 2025', then sort the table by Date descending"*
- *"Add data validation dropdowns to the Status column with options: Active, Pending, Completed"*
- *"Merge the header cells, center-align them, and auto-fit all column widths to content"*

The AI assistant analyzes your request, generates the proper Excel automation commands, and executes them **directly in your Excel application** - no formulas or programming knowledge required.

**Real-World Example - Power Query Optimization:**

```
You: "This Power Query is taking 5 minutes to refresh. Can you optimize it?"

AI Assistant (using ExcelMcp):
1. Exports the M code to analyze performance bottlenecks
2. Identifies queries that can use query folding
3. Refactors the M code to push operations to the data source
4. Updates the query in your workbook
5. Tests the refresh (now completes in 30 seconds)

Result: A professionally optimized Power Query with documented improvements
```

**🛡️ 100% Safe - Uses Excel's Native COM API**

Unlike third-party libraries that manipulate `.xlsx` files directly (risking file corruption), ExcelMcp uses **Excel's official COM API**. This ensures:
- ✅ **Zero risk of document corruption** - Excel handles all file operations safely
- ✅ **Interactive development** - See changes in real-time as you work with live Excel files
- ✅ **Timeout protection** - Built-in safeguards prevent indefinite hangs (2-min default, 5-min max)
- ✅ **Comprehensive automation** - Currently supports 166 operations across 11 specialized tools covering Power Query, Data Model/DAX, VBA, PivotTables, Excel Tables, ranges, and more

**🔧 Advanced Features:**
- **Batch Operations** - Group multiple operations in a single Excel session for 75-90% faster execution
- **Timeout Protection** - Automatic timeout enforcement with rich error messages and LLM-friendly guidance ([docs](docs/TIMEOUT-IMPLEMENTATION-GUIDE.md))
- **Progress Logging** - Real-time operation status updates via stderr (MCP protocol compatible)
- **Error Recovery** - Intelligent retry suggestions and operation-specific troubleshooting guidance

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
- 🔄 **Power Query** - 16 operations: create/edit M code, manage transformations, load configurations (worksheet, data model, connection only), error diagnostics, M code evaluation
- 📊 **Power Pivot (Data Model)** - 15 operations: build DAX measures, manage relationships, discover model structure (tables, columns), export to .dax files
- 🎨 **Excel Tables** - 26 operations: automate formatting, filtering, sorting, structured references, number formats, column management
- 📈 **PivotTables** - 20 operations: create and configure PivotTables for interactive analysis
- 📝 **VBA Macros** - 7 operations: export/import/run VBA code, integrate with version control
- 📋 **Ranges & Data** - 45 operations: values, formulas, copy/paste, find/replace, formatting, validation, merge, conditional formatting, cell protection
- 📄 **Worksheets** - 13 operations: lifecycle management, tab colors, visibility controls
- 🔌 **Connections** - 11 operations: manage OLEDB, ODBC, Text, Web data sources
- 🏷️ **Named Ranges** - 7 operations: named range management and bulk operations

<details>
<summary>📚 <strong>See Complete Feature List (100+ Operations)</strong></summary>

### Power Query & M Code (16 operations)
- Create, read, update, delete Power Query transformations
- Export/import M code for version control
- Manage query load destinations (worksheet/data model/connection-only/both)
- Set privacy levels for data source combinations
- Get load configuration for existing queries
- Error diagnostics and M code evaluation
- List Excel workbook sources for Power Query integration

### Data Model & DAX (Power Pivot) (15 operations)
- Create/update/delete DAX measures with format types (Currency, Percentage, Decimal, General)
- Manage table relationships (create, toggle active/inactive, delete)
- Discover model structure (tables, columns, measures, relationships)
- Export measures to .dax files for Git workflows
- Get comprehensive model information
- **Note:** DAX calculated columns are not supported (use Excel UI for calculated columns)

### Excel Tables (ListObjects) (26 operations)
- Lifecycle: create, resize, rename, delete, get info
- Styling: apply table styles, toggle totals row, set column totals
- Column management: add, remove, rename columns
- Data operations: append rows, apply filters (criteria/values), clear filters, get filter state
- Sorting: single-column sort, multi-column sort (up to 3 levels)
- Number formatting: get/set column number formats
- Advanced features: structured references, Data Model integration

### PivotTables (20 operations)
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
- **Data Validation** (3 actions): add validation rules (dropdowns, number/date/text rules), get validation info, remove validation
- **Hyperlinks** (4 actions): add, remove, list all, get specific hyperlink
- **Smart Range Operations** (3 actions): UsedRange, CurrentRegion, get range info (address, dimensions, format)
- **Merge Operations** (3 actions): merge cells, unmerge cells, get merge info
- **Auto-Sizing** (2 actions): auto-fit columns, auto-fit rows
- **Conditional Formatting** (2 actions): add conditional formatting, clear conditional formatting
- **Cell Protection** (2 actions): set cell lock status, get cell lock status
- **Formatting & Styling** (3 actions): get style, set style, format range
- **45 range operations total** covering all common Excel range manipulation needs
- **Worksheet management** (13 actions): lifecycle (create, rename, copy, delete), tab colors (set, get, clear), visibility controls (show, hide, very-hide, get/set status)

### Data Connections (11 operations)
- Manage OLEDB, ODBC, Text, Web connections
- Import/export connections via .odc files
- Update connection strings and properties
- Refresh connections and test connectivity
- Load connection-only connections to worksheet tables
- Get/set connection properties (refresh settings, background query, etc.)

### Named Ranges (7 operations)
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

1. **excel_powerquery** (16 actions) - Power Query M code: create, view, import, export, update, delete, manage load destinations, privacy levels, errors, eval, list Excel sources
2. **excel_datamodel** (15 actions) - Power Pivot (Data Model): CRUD DAX measures/relationships, discover structure (tables, columns), export to .dax files
3. **excel_table** (26 actions) - Excel Tables: lifecycle, columns, filters, sorts, structured references, totals, number formatting, Data Model integration
4. **excel_pivottable** (20 actions) - PivotTables: create from ranges/tables, field management (row/column/value/filter), aggregations, filters, sorting, extract data
5. **excel_range** (45 actions) - Ranges: get/set values/formulas, number formatting, visual formatting (font, fill, border, alignment), data validation, clear, copy, insert/delete, find/replace, sort, hyperlinks, merge, conditional formatting, cell protection
6. **excel_vba** (7 actions) - VBA: list, view, export, import, update, run, delete modules
7. **excel_connection** (11 actions) - Connections: OLEDB/ODBC/Text/Web management, properties, refresh, test
8. **excel_worksheet** (13 actions) - Worksheets: lifecycle (list, create, rename, copy, delete), tab colors (set-tab-color, get-tab-color, clear-tab-color), visibility (set-visibility, get-visibility, show, hide, very-hide)
9. **excel_namedrange** (7 actions) - Named ranges: list, get, set, create, create-bulk, delete, update
10. **excel_file** (3 actions) - File operations: create empty .xlsx/.xlsm workbooks, close workbook, test
11. **excel_batch** (3 actions) - Multi-operation performance: begin, commit, list

**Total: 11 tools with 166 actions**

> 📚 **[Complete MCP Server Guide →](src/ExcelMcp.McpServer/README.md)** - Detailed tool documentation and examples

---

## 📋 Additional Information

### Testing Philosophy

**Why No Unit Tests?** ExcelMcp uses integration tests exclusively because Excel COM cannot be meaningfully mocked. Our integration tests ARE our unit tests. See **[ADR-001: No Unit Tests](docs/ADR-001-NO-UNIT-TESTS.md)** for full architectural rationale.

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

`Excel Automation` • `Automate Excel with AI` • `MCP Server` • `Model Context Protocol` • `GitHub Copilot Excel` • `AI Excel Assistant` • `Power Query Automation` • `Power Query M Code` • `Power Pivot Automation` • `DAX Measures` • `DAX Automation` • `Data Model Automation` • `PivotTable Automation` • `VBA Automation` • `Excel Tables Automation` • `Excel AI Integration` • `COM Interop` • `Windows Excel Automation` • `Excel Development Tools` • `Excel Productivity` • `Excel Scripting` • `Conversational Excel` • `Natural Language Excel`
