# ExcelMcp - MCP Server for Microsoft Excel

[![Build MCP Server](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml)
[![Build CLI](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total)](https://github.com/sbroenne/mcp-server-excel/releases)
[![NuGet Downloads - MCP Server](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.McpServer.svg?label=MCP%20Installs)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)

[![NuGet MCP Server](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg?label=MCP%20Server)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![NuGet CLI](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.CLI.svg?label=CLI)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-8.0-blue.svg)](https://dotnet.microsoft.com/download/dotnet/8.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

**Automate Excel with AI - A Model Context Protocol (MCP) server for comprehensive Excel automation through conversational AI.**

**MCP Server for Excel** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands. Automate Power Query, DAX measures, VBA macros, PivotTables, Charts, formatting, and data transformations - no Excel programming knowledge required. 

**🛡️ 100% Safe - Uses Excel's Native COM API** - Zero risk of file corruption. Unlike third-party libraries that manipulate `.xlsx` files directly, this project uses Excel's official API ensuring complete safety and compatibility.

**💡 Interactive Development** - See results instantly in Excel. Create a query, run it, inspect the output, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing.

**Optional CLI Tool:** For advanced users who prefer command-line scripting, ExcelMcp includes a CLI interface for RPA workflows, CI/CD pipelines, and batch automation. Both interfaces share the same 172 operations.

## 🚀 Quick Start (1 Minute)

**Requirements:** Windows OS + Microsoft Excel 2016+

### ⭐ Recommended: VS Code Extension (One-Click Setup)

**Fastest way to get started - everything configured automatically: [Install from Marketplace](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)**

The extension opens automatically after installation with a quick start guide!

### For Visual Studio, Claude Desktop, Cursor, Windsurf, or other MCP clients:

📖 **[Complete Installation Guide →](docs/INSTALLATION.md)** - Step-by-step setup for all AI assistants with ready-to-use config files

**⚠️ Important:** Close all Excel files before using ExcelMcp. The server requires **exclusive access** to workbooks during automation (Excel COM limitation).

## 🎯 What You Can Do

**12 specialized tools with 172 operations:**

- 🔄 **Power Query** (9 ops) - Atomic workflows, M code management, load destinations
- 📊 **Data Model/DAX** (14 ops) - Measures, relationships, model structure
- 🎨 **Excel Tables** (24 ops) - Lifecycle, filtering, sorting, structured references
- 📈 **PivotTables** (25 ops) - Creation, fields, aggregations, data extraction
- 📉 **Charts** (14 ops) - Create, configure, manage series and formatting
- 📝 **VBA** (6 ops) - Modules, execution, version control
- 📋 **Ranges** (42 ops) - Values, formulas, formatting, validation, protection
- 📄 **Worksheets** (16 ops) - Lifecycle, colors, visibility, cross-workbook moves
- 🔌 **Connections** (9 ops) - OLEDB/ODBC management and refresh
- 🏷️ **Named Ranges** (6 ops) - Parameters and configuration
- 📁 **Files** (5 ops) - Session management and workbook creation
- 🎨 **Conditional Formatting** (2 ops) - Rules and clearing

📚 **[Complete Feature Reference →](FEATURES.md)** - Detailed documentation of all 172 operations


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


## 📋 Additional Information

### CLI for Direct Automation

xcelMcp includes a CLI interface for Excel automation without AI assistance. This is useful for RPA workflows, CI/CD pipelines, or batch processing scripts. Run `excelcli --help` for a categorized list of commands, or `excelcli sheet --help` (replace `sheet`) to view action-specific options. **Always follow the session pattern:** `excelcli session open <file>` → run commands with `--session <id>` → `excelcli session save/close <id>`. See **[CLI Guide](src/ExcelMcp.CLI/README.md)** for complete documentation.



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

> 📚 **[Complete MCP Server Guide →](src/ExcelMcp.McpServer/README.md)** - Detailed tool documentation and examples



## Project Information

**License:** MIT License - see [LICENSE](LICENSE) file

**Contributing:** See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for guidelines

**Built With:** This entire project was developed using GitHub Copilot AI assistance - mainly with Claude but lately with Auto-mode.

**Acknowledgments:**
- Microsoft Excel Team - For comprehensive COM automation APIs
- Model Context Protocol community - For the AI integration standard
- Open Source Community - For inspiration and best practices



### SEO & Discovery

`Excel Automation` • `Automate Excel with AI` • `MCP Server` • `Model Context Protocol` • `GitHub Copilot Excel` • `AI Excel Assistant` • `Power Query Automation` • `Power Query M Code` • `Power Pivot Automation` • `DAX Measures` • `DAX Automation` • `Data Model Automation` • `PivotTable Automation` • `VBA Automation` • `Excel Tables Automation` • `Excel AI Integration` • `COM Interop` • `Windows Excel Automation` • `Excel Development Tools` • `Excel Productivity` • `Excel Scripting` • `Conversational Excel` • `Natural Language Excel`
# test
