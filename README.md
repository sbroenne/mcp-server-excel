# ExcelMcp - MCP Server for Microsoft Excel

[![VS Code Marketplace Installs](https://img.shields.io/visual-studio-marketplace/i/sbroenne.excel-mcp?label=VS%20Code%20Installs)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=GitHub%20Downloads)](https://github.com/sbroenne/mcp-server-excel/releases)
[![NuGet Downloads - MCP Server](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.McpServer.svg?label=NuGet%20MCP%20Server)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![NuGet Downloads - CLI](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.CLI.svg?label=NuGet%20CLI)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)

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

**MCP Server for Excel** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands. Automate Power Query, DAX measures, VBA macros, PivotTables, Charts, formatting, and data transformations (23 tools with 214 operations).

**🛡️ 100% Safe - Uses Excel's Native COM API** - Zero risk of file corruption. Unlike third-party libraries that manipulate `.xlsx` files directly, this project uses Excel's official API ensuring complete safety and compatibility.

**💡 Interactive Development** - See results instantly in Excel. Create a query, run it, inspect the output, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing.

**🧪 LLM-Tested Quality** - Tool behavior validated with real LLM workflows using [pytest-aitest](https://github.com/sbroenne/pytest-aitest). We test that LLMs correctly understand and use our tools.

**Technical Requirements:**
- ⚠️ **Windows Only** - COM interop is Windows-specific
- ⚠️ **Excel Required** - Microsoft Excel 2016 or later must be installed
- ⚠️ **Desktop Environment** - Controls actual Excel process (not for server-side processing)

## 🎯 What You Can Do

**23 specialized tools with 214 operations:**

- 🔄 **Power Query** (1 tool, 11 ops) - Atomic workflows, M code management, load destinations
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
- 🧮 **Calculation Mode** (1 tool, 3 ops) - Get/set calculation mode and trigger recalculation
- 🎚️ **Slicers** (1 tool, 8 ops) - Interactive filtering for PivotTables and Tables
- 🎨 **Conditional Formatting** (1 tool, 2 ops) - Rules and clearing

📚 **[Complete Feature Reference →](FEATURES.md)** - Detailed documentation of all 214 operations


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


## 🚀 Quick Start

| Platform | Installation |
|----------|-------------|
| **VS Code** | [Install Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) (one-click, recommended) |
| **Claude Desktop** | Download `.mcpb` from [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest) |
| **Any MCP Client** | `dotnet tool install --global Sbroenne.ExcelMcp.McpServer` then `npx add-mcp "mcp-excel" --name excel-mcp` |
| **Details** | 📖 [Installation Guide](docs/INSTALLATION.md) |

**⚠️ Important:** Close all Excel files before using. The server requires exclusive access to workbooks during automation.


## 🔧 CLI vs MCP Server

This package provides both **CLI** and **MCP Server** interfaces. Choose based on your use case:

| Interface | Best For | Why |
|-----------|----------|-----|
| **CLI** (`excelcli`) | Coding agents (Copilot, Cursor, Windsurf) | **64% fewer tokens** - single tool, no large schemas. Auto-generated from Core code, ensuring 1:1 feature parity. |
| **MCP Server** | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent connection. Better for interactive, exploratory workflows. |

**⚡ CLI Commands:** Generated automatically from Core service definitions using Roslyn source generators. All 22 command categories maintain exact 1:1 parity with MCP tools through shared code generation. See [code generation docs](docs/DEVELOPMENT.md#-cli-command-code-generation) for details.

<details>
<summary>📊 Benchmark Results (same task, same model)</summary>

| Metric | CLI | MCP Server | Winner |
|--------|-----|------------|--------|
| **Tokens** | ~59K | ~163K | 🏆 CLI (64% fewer) |

**Key insight:** MCP sends 23 tool schemas to the LLM on each request (~100K+ tokens).

</details>

**Manual Installation:**
```powershell
# Step 1: Install the unified package (MCP Server + CLI)
dotnet tool install --global Sbroenne.ExcelMcp.McpServer

# Step 2: Auto-configure all your coding agents (requires Node.js)
npx add-mcp "mcp-excel" --name excel-mcp
```

> ⚠️ **Step 2 requires [Node.js](https://nodejs.org/)** for `npx`. Install with `winget install OpenJS.NodeJS.LTS` if needed.

```powershell
# Optional: Install agent skills for better AI guidance
npx skills add sbroenne/mcp-server-excel --skill excel-cli   # Coding agents
npx skills add sbroenne/mcp-server-excel --skill excel-mcp   # Conversational AI
```

> 💡 **Skills provide AI guidance** - The CLI skill is highly recommended (agents don't work perfectly with CLI without it). The MCP skill is recommended - it adds workflow best practices and reduces token usage.


## ⚙️ How It Works - COM Automation & Unified Service Architecture

**ExcelMcp uses Windows COM automation to control the actual Excel application (not just .xlsx files).**

Both the **MCP Server** and **CLI** communicate with a shared **ExcelMCP Service** that manages Excel sessions. This unified architecture enables:

```
┌─────────────────────┐     ┌─────────────────────┐
│   MCP Server        │     │   CLI (excelcli)    │
│  (AI assistants)    │     │  (coding agents)    │
└─────────┬───────────┘     └─────────┬───────────┘
          │                           │
          └──────────┬────────────────┘
                     ▼
          ┌─────────────────────────┐
          │   ExcelMCP Service      │
          │  (shared session mgmt)  │
          └─────────┬───────────────┘
                    ▼
          ┌─────────────────────────┐
          │   Excel COM API         │
          │  (Excel.Application)    │
          └─────────────────────────┘
```

**Key Benefits:**
- ✅ **Shared Sessions** - CLI and MCP Server can access the same open workbooks
- ✅ **Single Excel Instance** - No duplicate Excel processes or file locks
- ✅ **System Tray UI** - Monitor active sessions via the ExcelMCP tray icon

**💡 Tip: Watch Excel While AI Works**
By default, Excel runs hidden for faster automation. To see changes in real-time, just ask:
- *"Show me Excel while you work"*
- *"Let me watch what you're doing"*
- *"Open Excel so I can see the changes"*

The AI will display the Excel window so you can watch every operation happen live - great for learning or verifying changes!

## 📋 Additional Information

📚 **[CLI Guide →](src/ExcelMcp.CLI/README.md)** | **[CLI Skill for Agents →](skills/excel-cli/SKILL.md)** | **[MCP Server Guide →](src/ExcelMcp.McpServer/README.md)** | **[All Agent Skills →](skills/README.md)**

**License:** MIT License - see [LICENSE](LICENSE) file

**Privacy:** See [PRIVACY.md](PRIVACY.md) for our privacy policy

**Contributing:** See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for guidelines

**Built With:** This entire project was developed using GitHub Copilot AI assistance - mainly with Claude but lately with Auto-mode.

**Acknowledgments:**
- Microsoft Excel Team - For comprehensive COM automation APIs
- Model Context Protocol community - For the AI integration standard
- Open Source Community - For inspiration and best practices

## Related Projects

Other projects by the author:

- [pytest-aitest](https://github.com/sbroenne/pytest-aitest) — LLM-powered testing framework for AI agents
- [Windows MCP Server](https://windowsmcpserver.dev/) — AI-powered Windows automation via MCP
- [OBS Studio MCP Server](https://github.com/sbroenne/mcp-server-obs) — AI-powered OBS Studio automation
- [HeyGen MCP Server](https://github.com/sbroenne/heygen-mcp) — MCP server for HeyGen AI video generation
