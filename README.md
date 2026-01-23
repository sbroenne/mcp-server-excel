# ExcelMcp - MCP Server for Microsoft Excel

[![VS Code Marketplace Installs](https://img.shields.io/visual-studio-marketplace/i/sbroenne.excel-mcp?label=VS%20Code%20Installs)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=GitHub%20Downloads)](https://github.com/sbroenne/mcp-server-excel/releases)
[![NuGet Downloads - MCP Server](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.McpServer.svg?label=Nuget%20MCP%20Server%20Installs)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)

[![Build MCP Server](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml)
[![Build CLI](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![NuGet MCP Server](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg?label=MCP%20Server)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![NuGet CLI](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.CLI.svg?label=CLI)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-10-blue.svg)](https://dotnet.microsoft.com/download/dotnet/10.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

**Automate Excel with AI - A Model Context Protocol (MCP) server for comprehensive Excel automation through conversational AI.**

**MCP Server for Excel** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands. Automate Power Query, DAX measures, VBA macros, PivotTables, Charts, formatting, and data transformations (22 tools with 206 operations).

**🛡️ 100% Safe - Uses Excel's Native COM API** - Zero risk of file corruption. Unlike third-party libraries that manipulate `.xlsx` files directly, this project uses Excel's official API ensuring complete safety and compatibility.

**💡 Interactive Development** - See results instantly in Excel. Create a query, run it, inspect the output, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing.

**🧪 LLM-Tested Quality** - Tool behavior validated with real AI agents using [agent-benchmark](https://github.com/mykhaliev/agent-benchmark). We test that LLMs correctly understand and use our tools.

**Optional CLI Tool:** For advanced users who prefer command-line scripting, ExcelMcp includes a CLI interface for RPA workflows, CI/CD pipelines, and batch automation. 

## 🚀 Quick Start (1 Minute)

**Requirements:** Windows OS + Microsoft Excel 2016+

### ⭐ Recommended: VS Code Extension (One-Click Setup)

**Fastest way to get started - everything configured automatically: [Install from Marketplace](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)**

The extension opens automatically after installation with a quick start guide!

### For Claude Desktop (One-Click Install)

Download the `.mcpb` file from the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest) and double-click to install.

### For Visual Studio, Cursor, Windsurf, or other MCP clients:

📖 **[Complete Installation Guide →](docs/INSTALLATION.md)** - Step-by-step setup for all AI assistants with ready-to-use config files

**⚠️ Important:** Close all Excel files before using ExcelMcp. The server requires **exclusive access** to workbooks during automation (Excel COM limitation).

## 🎯 What You Can Do

**22 specialized tools with 206 operations:**

- 🔄 **Power Query** (1 tool, 10 ops) - Atomic workflows, M code management, load destinations
- 📊 **Data Model/DAX** (2 tools, 18 ops) - Measures with auto-formatted DAX, relationships, model structure
- 🎨 **Excel Tables** (2 tools, 27 ops) - Lifecycle, filtering, sorting, structured references
- 📈 **PivotTables** (3 tools, 30 ops) - Creation, fields, aggregations, calculated members/fields
- 📉 **Charts** (2 tools, 26 ops) - Create, configure, series, formatting, data labels, trendlines
- 📝 **VBA** (1 tool, 6 ops) - Modules, execution, version control
- 📋 **Ranges** (4 tools, 42 ops) - Values, formulas, formatting, validation, protection
- 📄 **Worksheets** (2 tools, 16 ops) - Lifecycle, colors, visibility, cross-workbook moves
- 🔌 **Connections** (1 tool, 9 ops) - OLEDB/ODBC management and refresh
- 🏷️ **Named Ranges** (1 tool, 6 ops) - Parameters and configuration
- 📁 **Files** (1 tool, 6 ops) - Session management and workbook creation
- �️ **Slicers** (1 tool, 8 ops) - Interactive filtering for PivotTables and Tables
- �🎨 **Conditional Formatting** (1 tool, 2 ops) - Rules and clearing

📚 **[Complete Feature Reference →](FEATURES.md)** - Detailed documentation of all 206 operations


## 💬 Example Prompts

**Create & Populate Data:**
- *"Create a new Excel file called SalesTracker.xlsx with a table for Date, Product, Quantity, Unit Price, and Total with sample data"*
- *"Put this data in A1:C4 - Name, Age, City / Alice, 30, Seattle / Bob, 25, Portland"*
- *"Add a formula column that calculates Quantity times Unit Price"*

**Analysis & Visualization:**
- *"Create a PivotTable from this data showing total sales by Product, then add a bar chart"*
- *"Use Power Query to import products.csv, load it to the Data Model, and create a measure for Total Revenue"*
- *"Create a slicer for the Region field so I can filter the PivotTable interactively"*
- *"Create a relationship between the Orders and Products tables using ProductID"*

**Formatting & Styling:**
- *"Format the Price column as currency and highlight values over $500 in green"*
- *"Convert this range to an Excel Table with a blue style and add a totals row"*
- *"Make the headers bold with a dark background and auto-fit column widths"*

**Automation:**
- *"Export all Power Query M code to files for version control"*
- *"Run the UpdatePrices macro"*
- *"Show me Excel while you work"* - watch changes in real-time

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

ExcelMcp includes a CLI interface for Excel automation without AI assistance. This is useful for RPA workflows, CI/CD pipelines, or batch processing scripts. Run `excelcli --help` for a categorized list of commands, or `excelcli sheet --help` (replace `sheet`) to view action-specific options. **Always follow the session pattern:** `excelcli session open <file>` → run commands with `--session <id>` → `excelcli session save/close <id>`. See **[CLI Guide](src/ExcelMcp.CLI/README.md)** for complete documentation.

### Agent Skills (Cross-Platform AI Guidance)

ExcelMcp includes an **Agent Skills** package for cross-platform AI assistant guidance. Skills enable GitHub Copilot, Claude Code, Cursor, Windsurf, Gemini CLI, and other agents to effectively use Excel MCP Server.

| Platform | Installation |
|----------|--------------|
| **GitHub Copilot** | Auto-installed by VS Code extension to `~/.copilot/skills/` |
| **Claude Code** | Copy [CLAUDE.md](skills/CLAUDE.md) to project root |
| **Cursor** | Copy [.cursorrules](skills/.cursorrules) to project root |
| **Others** | `npx add-skill sbroenne/mcp-server-excel` or download from releases |

📚 **[Agent Skills →](skills/README.md)** | 📄 **[SKILL.md →](skills/excel-mcp/SKILL.md)** | ⚙️ **[CLAUDE.md →](skills/CLAUDE.md)**

For VS Code Copilot, enable the setting `chat.useAgentSkills` to load skills.



## 🔧 How It Works - COM Interop Architecture

**ExcelMcp uses Windows COM automation to control the actual Excel application (not just .xlsx files).**

This means you get:
- ✅ **Full Excel Feature Access** - Power Query engine, VBA runtime, Data Model, calculation engine, pivot tables
- ✅ **True Compatibility** - Works exactly like Excel UI, no feature limitations
- ✅ **Live Data Operations** - Refresh Power Query, connections, Data Model in real workbooks
- ✅ **Interactive Development** - Immediate Excel feedback as AI makes changes
- ✅ **All File Formats** - .xlsx, .xlsm, .xlsb, even legacy formats

**💡 Tip: Watch Excel While AI Works**
By default, Excel runs hidden for faster automation. To see changes in real-time, just ask:
- *"Show me Excel while you work"*
- *"Let me watch what you're doing"*
- *"Open Excel so I can see the changes"*

The AI will display the Excel window so you can watch every operation happen live - great for learning or verifying changes!

**Technical Requirements:**
- ⚠️ **Windows Only** - COM interop is Windows-specific
- ⚠️ **Excel Required** - Microsoft Excel 2016 or later must be installed
- ⚠️ **Desktop Environment** - Controls actual Excel process (not for server-side processing)

> 📚 **[Complete MCP Server Guide →](src/ExcelMcp.McpServer/README.md)** - Detailed tool documentation and examples



## Project Information

**License:** MIT License - see [LICENSE](LICENSE) file

**Privacy:** See [PRIVACY.md](PRIVACY.md) for our privacy policy

**Contributing:** See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for guidelines

**Built With:** This entire project was developed using GitHub Copilot AI assistance - mainly with Claude but lately with Auto-mode.

**Acknowledgments:**
- Microsoft Excel Team - For comprehensive COM automation APIs
- Model Context Protocol community - For the AI integration standard
- Open Source Community - For inspiration and best practices
