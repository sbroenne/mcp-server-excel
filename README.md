# ExcelMcp - MCP Server for Microsoft Excel

[![VS Code Marketplace Installs](https://img.shields.io/visual-studio-marketplace/i/sbroenne.excel-mcp?label=VS%20Code%20Installs)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=GitHub%20Downloads)](https://github.com/sbroenne/mcp-server-excel/releases)

[![Build MCP Server](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml)
[![Build CLI](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-10-blue.svg)](https://dotnet.microsoft.com/download/dotnet/10.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

**Automate Excel with AI - A Model Context Protocol (MCP) server for comprehensive Excel automation through conversational AI.**

**MCP Server for Excel** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands. Automate Power Query, DAX measures, VBA macros, PivotTables, Charts, formatting, and data transformations (25 tools with 230 operations).

**🛡️ 100% Safe - Uses Excel's Native COM API** - Zero risk of file corruption. Unlike third-party libraries that manipulate `.xlsx` files directly, this project uses Excel's official API ensuring complete safety and compatibility.

**💡 Interactive Development** - See results instantly in Excel. Create a query, run it, inspect the output, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing.

**🧪 LLM-Tested Quality** - Tool behavior validated with real LLM workflows using [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering). We test that LLMs correctly understand and use our tools.

**Technical Requirements:**
- ⚠️ **Windows Only** - COM interop is Windows-specific
- ⚠️ **Excel Required** - Microsoft Excel 2016 or later must be installed
- ⚠️ **Desktop Environment** - Controls actual Excel process (not for server-side processing)

## 🎯 What You Can Do

**25 specialized tools with 230 operations:**

- 🔄 **Power Query** (1 tool, 12 ops) - Atomic workflows, M code management, load destinations
- 📊 **Data Model/DAX** (2 tools, 19 ops) - Measures with auto-formatted DAX, relationships, model structure
- 🎨 **Excel Tables** (2 tools, 27 ops) - Lifecycle, filtering, sorting, structured references
- 📈 **PivotTables** (3 tools, 30 ops) - Creation, fields, aggregations, calculated members/fields
- 📉 **Charts** (2 tools, 29 ops) - Create, configure, series, formatting, data labels, trendlines
- 📝 **VBA** (1 tool, 6 ops) - Modules, execution, version control
- 📋 **Ranges** (4 tools, 46 ops) - Values, formulas, formatting, validation, protection
- 📄 **Worksheets** (2 tools, 16 ops) - Lifecycle, colors, visibility, cross-workbook moves
- 🔌 **Connections** (1 tool, 9 ops) - OLEDB/ODBC management and refresh
- 🏷️ **Named Ranges** (1 tool, 6 ops) - Parameters and configuration
- 📁 **Files** (1 tool, 6 ops) - Session management, workbook creation, IRM/AIP-protected file support
- 🧮 **Calculation Mode** (1 tool, 3 ops) - Get/set calculation mode and trigger recalculation
- 🎚️ **Slicers** (1 tool, 8 ops) - Interactive filtering for PivotTables and Tables
- 🎨 **Conditional Formatting** (1 tool, 2 ops) - Rules and clearing
- 📸 **Screenshot** (1 tool, 2 ops) - Capture ranges/sheets as PNG for LLM visual verification
- 🪧 **Window Management** (1 tool, 9 ops) - Show/hide Excel, arrange, position, status bar feedback

📚 **[Complete Feature Reference →](FEATURES.md)** - Detailed documentation of all 230 operations


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
- *"Apply the same section-header styling to A1:G1, A12:G12, and A24:G24 in one step"*

Formatting split: number display formats use the `range` tool, while visual styling and auto-fit use `range_format`.

**Automation:**
- *"Export all Power Query M code to files for version control"*
- *"Run the UpdatePrices macro"*
- *"Show me Excel while you work"* - watch changes in real-time

**🪟 Agent Mode — Watch AI Work in Excel:**
- *"Show me Excel side-by-side while you build this dashboard"* - real-time visibility
- *"Let me watch while you create the chart"* - AI asks your preference, then shows Excel
- Status bar shows live progress: *"ExcelMcp: Building PivotTable from Sales data..."*

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
| **Any MCP Client** | Download `mcp-excel.exe` from [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest) and add to PATH |
| **Details** | 📖 [Installation Guide](docs/INSTALLATION.md) |

**⚠️ Important:** Close all Excel files before using. The server requires exclusive access to workbooks during automation.


## 🔧 CLI vs MCP Server

This package provides both **CLI** and **MCP Server** interfaces. Choose based on your use case:

| Interface | Best For | Why |
|-----------|----------|-----|
| **CLI** (`excelcli`) | Coding agents (Copilot, Cursor, Windsurf) + Scripting | **64% fewer tokens** - single tool, no large schemas. Auto-generated from Core code, ensuring 1:1 feature parity. Bundled with excel-cli skill. |
| **MCP Server** | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent connection. Better for interactive, exploratory workflows. |

**Installation:**
- **CLI via Copilot plugin** (Recommended for Copilot CLI): Install the `excel-cli` plugin — bundles `excelcli.exe` and the `excel-cli` skill together
- **CLI Standalone**: Download ZIP from [releases](https://github.com/sbroenne/mcp-server-excel/releases/latest) or install via NuGet
- **Skill only**: Install the `excel-cli` skill separately when your agent already has `excelcli` available on PATH
- **MCP Server**: Download from releases or install VS Code Extension

**⚡ CLI Commands:** Generated automatically from Core service definitions using Roslyn source generators. All 22 command categories maintain exact 1:1 parity with MCP tools through shared code generation. See [code generation docs](docs/DEVELOPMENT.md#-cli-command-code-generation) for details.

### 📦 GitHub Copilot Plugins (Alternative Installation)

ExcelMcp is available as two distributable **GitHub Copilot plugins** published through the GitHub Copilot plugin marketplace:

```powershell
# Register the plugin marketplace (one-time)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install one or both plugins with Copilot CLI
copilot plugin install excel-mcp@mcp-server-excel-plugins   # For conversational AI
copilot plugin install excel-cli@mcp-server-excel-plugins   # For scripting / coding agents (bundles excelcli.exe)
```

- **`excel-mcp`** — MCP-server-centric workflows
- **`excel-cli`** — token-efficient CLI workflows with bundled self-contained `excelcli.exe` (no .NET runtime required)
- Install **either plugin or both**, depending on your agent surface and workflow

The published repo [`sbroenne/mcp-server-excel-plugins`](https://github.com/sbroenne/mcp-server-excel-plugins) is the actual Copilot CLI marketplace. This source repo is **not** itself a marketplace; `.github/plugins/` only contains source-owned overlay files that the publish workflow copies into the published plugin directories.

These are **GitHub Copilot marketplace packages**, not a generic cross-tool install command. The commands above are the documented Copilot CLI install path. VS Code also supports agent plugins in preview, and Claude has its own plugin system, but those surfaces have their own installation and enablement flows.

These plugins are republished automatically after each successful ExcelMcp release by a follow-on workflow that uses a stored cross-repo token scoped to the published marketplace repo. That publish path is sync-gated (no downstream republish when plugin-facing install artifacts did not change), keeps downgrade/tag mismatches blocked, and still exposes a manual maintainer re-sync path for repair/replay scenarios.

After installing `excel-cli`, run the bundled one-time helper to expose the plugin-shipped CLI on PATH:

```powershell
pwsh -ExecutionPolicy Bypass -File "$env:USERPROFILE\.copilot\installed-plugins\mcp-server-excel-plugins\excel-cli\bin\install-global.ps1"
```

**📖 [Plugin Installation Guide →](docs/INSTALLATION.md#github-copilot-plugins-alternative-installation)** | **[Published Marketplace Repo →](https://github.com/sbroenne/mcp-server-excel-plugins)** | **[VS Code Agent Plugins →](https://code.visualstudio.com/docs/copilot/customization/agent-plugins)** | **[Claude Plugins Reference →](https://code.claude.com/docs/en/plugins-reference)**

<details>
<summary>📊 Benchmark Results (same task, same model)</summary>

| Metric | CLI | MCP Server | Winner |
|--------|-----|------------|--------|
| **Tokens** | ~59K | ~163K | 🏆 CLI (64% fewer) |

**Key insight:** MCP sends 23 tool schemas to the LLM on each request (~100K+ tokens).

</details>

**Manual Installation:**
```powershell
# Primary: Download standalone executables from latest release (no .NET runtime required)
# https://github.com/sbroenne/mcp-server-excel/releases/latest
# - ExcelMcp-MCP-Server-{version}-windows.zip → extract mcp-excel.exe
# - ExcelMcp-CLI-{version}-windows.zip → extract excelcli.exe (optional, for scripting)

# Secondary: Install via .NET tool (requires .NET 10 runtime)
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
dotnet tool install --global Sbroenne.ExcelMcp.CLI

# After installing either way, auto-configure all your coding agents:
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

**Releasing:** See [RELEASE-STRATEGY.md](docs/RELEASE-STRATEGY.md) for the unified release workflow (MCP Server, CLI, VS Code Extension, MCPB, Agent Skills, and GitHub Copilot plugins, including the cross-repo PAT-backed plugin republish flow)

**Contributing:** See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for guidelines

**Built With:** This entire project was developed using GitHub Copilot AI assistance - mainly with Claude but lately with Auto-mode.

**Acknowledgments:**
- Microsoft Excel Team - For comprehensive COM automation APIs
- Model Context Protocol community - For the AI integration standard
- Open Source Community - For inspiration and best practices

## Related Projects

Other projects by the author:

- [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering) — LLM-powered testing framework for AI agents
- [Windows MCP Server](https://windowsmcpserver.dev/) — AI-powered Windows automation via MCP
- [OBS Studio MCP Server](https://github.com/sbroenne/mcp-server-obs) — AI-powered OBS Studio automation
- [HeyGen MCP Server](https://github.com/sbroenne/heygen-mcp) — MCP server for HeyGen AI video generation
