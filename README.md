# ExcelMcp - MCP Server for Microsoft Excel [![VS Code Marketplace Installs](https://img.shields.io/visual-studio-marketplace/i/sbroenne.excel-mcp?label=VS%20Code%20Installs)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=GitHub%20Downloads)](https://github.com/sbroenne/mcp-server-excel/releases)
[![NuGet Downloads - MCP Server](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.McpServer.svg?label=NuGet%20MCP%20Server)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![NuGet Downloads - CLI](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.CLI.svg?label=NuGet%20CLI)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI) [![Build MCP Server](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml)
[![Build CLI](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![NuGet MCP Server](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg?label=MCP%20Server)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![NuGet CLI](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.CLI.svg?label=CLI)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI) [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-10-blue.svg)](https://dotnet.microsoft.com/download/dotnet/10.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/) **Automate Excel with AI - A Model Context Protocol (MCP) server for comprehensive Excel automation through conversational AI.** **MCP Server for Excel** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands. Automate Power Query, DAX measures, VBA macros, PivotTables, Charts, formatting, and data transformations (22 tools with 211 operations). ### CLI vs MCP Server This package provides both **CLI** and **MCP Server** interfaces. Choose based on your use case: | Interface | Best For | Why |
|-----------|----------|-----|
| **CLI** (`excelcli`) | Coding agents (Copilot, Cursor, Windsurf) | **64% fewer tokens** - single tool, no large schemas. Better for cost-sensitive, high-throughput automation. |
| **MCP Server** | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent connection. Better for interactive, exploratory workflows. | <details>
<summary>📊 Benchmark Results (same task, same model)</summary> | Metric | CLI | MCP Server | Winner |
|--------|-----|------------|--------|
| **Tokens** | ~59K | ~163K | 🏆 CLI (64% fewer) | **Key insight:** MCP sends 22 tool schemas to the LLM on each request (~100K+ tokens). CLI wraps everything in one `excel_execute` tool and offloads guidance to a skill file. </details> **Installation:**
```powershell
# CLI for coding agents
dotnet tool install --global Sbroenne.ExcelMcp.CLI
excelcli --help # MCP Server for AI assistants (or use VS Code extension)
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
``` **🛡️ 100% Safe - Uses Excel's Native COM API** - Zero risk of file corruption. Unlike third-party libraries that manipulate `.xlsx` files directly, this project uses Excel's official API ensuring complete safety and compatibility. **TIP: Interactive Development** - See results instantly in Excel. Create a query, run it, inspect the output, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing. **🧪 LLM-Tested Quality** - Tool behavior validated with real AI agents using [agent-benchmark](https://github.com/mykhaliev/agent-benchmark). We test that LLMs correctly understand and use our tools. **Optional CLI Tool:** For advanced users who prefer command-line scripting, ExcelMcp includes a CLI interface for RPA workflows, CI/CD pipelines, and batch automation. ## Quick Start (1 Minute) **Requirements:** Windows OS + Microsoft Excel 2016+ ### RECOMMENDED: Recommended: VS Code Extension (One-Click Setup) **Fastest way to get started - everything configured automatically: [Install from Marketplace](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)** The extension opens automatically after installation with a quick start guide! ### For Claude Desktop (One-Click Install) Download the `.mcpb` file from the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest) and double-click to install. ### For Visual Studio, Cursor, Windsurf, or other MCP clients: 📖 **[Complete Installation Guide →](docs/INSTALLATION.md)** - Step-by-step setup for all AI assistants with ready-to-use config files **WARNING: Important:** Close all Excel files before using ExcelMcp. The server requires **exclusive access** to workbooks during automation (Excel COM limitation). ## What You Can Do **22 specialized tools with 211 operations:** - 🔄 **Power Query** (1 tool, 11 ops) - Atomic workflows, M code management, load destinations
- 📊 **Data Model/DAX** (2 tools, 18 ops) - Measures with auto-formatted DAX, relationships, model structure
- 🎨 **Excel Tables** (2 tools, 27 ops) - Lifecycle, filtering, sorting, structured references
- 📈 **PivotTables** (3 tools, 30 ops) - Creation, fields, aggregations, calculated members/fields
- 📉 **Charts** (2 tools, 26 ops) - Create, configure, series, formatting, data labels, trendlines
- **VBA** (1 tool, 6 ops) - Modules, execution, version control
- 📋 **Ranges** (4 tools, 42 ops) - Values, formulas, formatting, validation, protection
- 📄 **Worksheets** (2 tools, 16 ops) - Lifecycle, colors, visibility, cross-workbook moves
- 🔌 **Connections** (1 tool, 9 ops) - OLEDB/ODBC management and refresh
- 🏷️ **Named Ranges** (1 tool, 6 ops) - Parameters and configuration
- 📁 **Files** (1 tool, 6 ops) - Session management and workbook creation
- �️ **Slicers** (1 tool, 8 ops) - Interactive filtering for PivotTables and Tables
- �🎨 **Conditional Formatting** (1 tool, 2 ops) - Rules and clearing DOCS: **[Complete Feature Reference →](FEATURES.md)** - Detailed documentation of all 211 operations ## 💬 Example Prompts **Create & Populate Data:**
- *"Create a new Excel file called SalesTracker.xlsx with a table for Date, Product, Quantity, Unit Price, and Total with sample data"*
- *"Put this data in A1:C4 - Name, Age, City / Alice, 30, Seattle / Bob, 25, Portland"*
- *"Add a formula column that calculates Quantity times Unit Price"* **Analysis & Visualization:**
- *"Create a PivotTable from this data showing total sales by Product, then add a bar chart"*
- *"Use Power Query to import products.csv, load it to the Data Model, and create a measure for Total Revenue"*
- *"Create a slicer for the Region field so I can filter the PivotTable interactively"*
- *"Create a relationship between the Orders and Products tables using ProductID"* **Formatting & Styling:**
- *"Format the Price column as currency and highlight values over $500 in green"*
- *"Convert this range to an Excel Table with a blue style and add a totals row"*
- *"Make the headers bold with a dark background and auto-fit column widths"* **Automation:**
- *"Export all Power Query M code to files for version control"*
- *"Run the UpdatePrices macro"*
- *"Show me Excel while you work"* - watch changes in real-time ## 👥 Who Should Use This? **Perfect for:**
- YES: **Data analysts** automating repetitive Excel workflows
- YES: **Developers** building Excel-based data solutions
- YES: **Business users** managing complex Excel workbooks
- YES: **Teams** maintaining Power Query/VBA/DAX code in Git **Not suitable for:**
- NO: Server-side data processing (use libraries like ClosedXML, EPPlus instead)
- NO: Linux/macOS users (Windows + Excel installation required)
- NO: High-volume batch operations (consider Excel-free alternatives) ## 📋 Additional Information ### CLI for Coding Agents (Recommended) **For coding agents like GitHub Copilot, Cursor, and Windsurf, use the CLI instead of MCP Server.** CLI invocations are more token-efficient: they avoid loading large tool schemas into the model context, allowing agents to act through concise commands. ```powershell
# Install CLI
dotnet tool install --global Sbroenne.ExcelMcp.CLI # Agent workflow (use -q for clean JSON output)
excelcli -q session open C:\Data\Report.xlsx # Returns {"sessionId":1,...}
excelcli -q range set-values --session 1 --sheet Sheet1 --range A1 --values '[["Hello"]]'
excelcli -q session save --session 1
excelcli -q session close --session 1
``` **Key features:**
- `-q` / `--quiet` flag for clean JSON output (no banner)
- Auto-suppresses banner when output is piped
- All commands output parseable JSON
- Session pattern for efficient Excel reuse Run `excelcli --help` for all commands, or `excelcli <command> --help` for action-specific options. DOCS: **[CLI Skill for Agents →](skills/excel-cli/SKILL.md)** | **[CLI Guide →](src/ExcelMcp.CLI/README.md)** ### Agent Skills (Cross-Platform AI Guidance) Skills are auto-installed by the VS Code extension. For other platforms: ```powershell
npx add-skill sbroenne/mcp-server-excel
``` DOCS: **[Agent Skills →](skills/README.md)** ## How It Works - COM Interop Architecture **ExcelMcp uses Windows COM automation to control the actual Excel application (not just .xlsx files).** This means you get:
- YES: **Full Excel Feature Access** - Power Query engine, VBA runtime, Data Model, calculation engine, pivot tables
- YES: **True Compatibility** - Works exactly like Excel UI, no feature limitations
- YES: **Live Data Operations** - Refresh Power Query, connections, Data Model in real workbooks
- YES: **Interactive Development** - Immediate Excel feedback as AI makes changes
- YES: **All File Formats** - .xlsx, .xlsm, .xlsb, even legacy formats **TIP: Tip: Watch Excel While AI Works**
By default, Excel runs hidden for faster automation. To see changes in real-time, just ask:
- *"Show me Excel while you work"*
- *"Let me watch what you're doing"*
- *"Open Excel so I can see the changes"* The AI will display the Excel window so you can watch every operation happen live - great for learning or verifying changes! **Technical Requirements:**
- WARNING: **Windows Only** - COM interop is Windows-specific
- WARNING: **Excel Required** - Microsoft Excel 2016 or later must be installed
- WARNING: **Desktop Environment** - Controls actual Excel process (not for server-side processing) > DOCS: **[Complete MCP Server Guide →](src/ExcelMcp.McpServer/README.md)** - Detailed tool documentation and examples ## Project Information **License:** MIT License - see [LICENSE](LICENSE) file **Privacy:** See [PRIVACY.md](PRIVACY.md) for our privacy policy **Contributing:** See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for guidelines **Built With:** This entire project was developed using GitHub Copilot AI assistance - mainly with Claude but lately with Auto-mode. **Acknowledgments:**
- Microsoft Excel Team - For comprehensive COM automation APIs
- Model Context Protocol community - For the AI integration standard
- Open Source Community - For inspiration and best practices
