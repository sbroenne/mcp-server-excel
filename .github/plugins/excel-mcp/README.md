# Excel MCP Plugin

**Model Context Protocol server for natural language Excel automation**

This plugin provides the `excel-mcp` skill and a plugin-local MCP bootstrap for GitHub Copilot. Use natural language to automate Power Query, DAX measures, PivotTables, Tables, Charts, VBA macros, and more through Windows Excel COM API.

**Best for:** Conversational AI workflows (GitHub Copilot Chat, Claude Desktop, Cursor) where rich tool schemas and persistent connections matter more than token efficiency.

---

## Prerequisites

- **Windows** with Microsoft Excel 2016 or later (COM interop required)
- **GitHub Copilot extension** or other MCP-compatible client

---

## Installation

### Option 1: VS Code Extension (Recommended)

Install the [Excel MCP VS Code extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) for one-click setup with GitHub Copilot.

### Option 2: Plugin Marketplace

```powershell
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
copilot plugin install excel-mcp@mcp-server-excel-plugins
```

### Option 3: Manual Installation

1. Install the plugin
2. Let the plugin bootstrap the latest self-contained `mcp-excel.exe` on first use
3. Or add the standalone binary to your MCP client configuration manually (see [MCP Server Installation Guide](https://excelmcpserver.dev/installation-mcp-server/))

### Runtime Bootstrap

The plugin does **not** rely on a bundled `mcp-excel.exe`. Its Agent Plugins 1.0 `mcp.json` launches a PowerShell wrapper that:

- checks GitHub Releases for the newest `ExcelMcp-MCP-Server-*-windows.zip`
- downloads and caches the latest self-contained Windows server on first invocation
- stores plugin-hosted runtime state under `PLUGIN_DATA\runtime`
- re-checks freshness at most once per Copilot chat session

The optional global shim runs outside an Agent Plugins host, uses
`~\.copilot\plugin-runtime\mcp-server-excel\excel-mcp`, and checks for updates at
most once every 24 hours.

If you want the server registered globally in `~/.copilot/mcp-config.json`, run:

```powershell
pwsh -ExecutionPolicy Bypass -File `
  "$env:USERPROFILE\.copilot\installed-plugins\mcp-server-excel-plugins\excel-mcp\com.github.copilot\bin\install-global.ps1"
```

---

## What You Can Do

**31 specialized tools with 326 operations** for comprehensive Excel automation:

### Core Operations

- **Power Query** (12 ops) — Create, update, refresh; optional remote M code formatting
- **Data Model/DAX** (20 ops) — Measures, relationships, source metadata, EVALUATE queries
- **PivotTables** (35 ops) — Fields, grouping, cache options, drill-through, calculations
- **Excel Tables** (27 ops) — Lifecycle, filtering, sorting, DAX-backed tables
- **Charts** (33 ops) — Combo series, plotting, placement, formatting, labels, trendlines
- **Ranges** (51 ops) — Values, formulas, hyperlinks, threaded comments, formatting, validation
- **Worksheets** (33 ops) — Lifecycle, outlines, protection, notes, images, shapes, page setup
- **Workbooks** (15 ops) — Metadata, properties, Save As/copy, PDF/XPS, external links

### Advanced Features

- **VBA** (6 ops) — Module import/export, run procedures, version control
- **Connections** (11 ops) — OLEDB/ODBC connections with refresh status and cancellation
- **QueryTables** (9 ops) — Text/CSV and legacy HTML imports
- **Drawing Objects** (14 ops) — Shapes, Forms controls, and sparklines
- **What-If Analysis** (8 ops) — Goal Seek, scenarios, summaries, data tables
- **XML Maps** (6 ops) — Schemas, XPath mapping, secure import/export
- **Slicers** (8 ops) — Interactive filtering for PivotTables and Tables
- **Conditional Formatting** (4 ops) — Add and clear rules; list range or worksheet rules
- **Named Ranges** (6 ops) — Create, update, delete named ranges
- **Calculation Mode** (3 ops) — Get/set mode, trigger recalculation (performance optimization)
- **Python in Excel** (2 ops) — Set/get `=PY()` formulas and results
- **Screenshot** (2 ops) — Capture ranges/sheets as PNG for visual verification
- **Window Management** (15 ops) — Show/hide, panes, zoom, display options, positioning

### File Operations

- **File Operations** (5 ops) — Create, open, close, list, and test files
- **IRM/AIP Support** — Auto-detects protected files, opens with Excel visible for authentication

**Complete documentation:** [Full Feature Reference](https://excelmcpserver.dev/features/)

---

## Why MCP Server?

| Interface | Best For | Key Benefit |
|-----------|----------|-------------|
| **MCP Server** | Conversational AI (Claude, Copilot Chat) | Rich tool discovery, persistent sessions |
| **CLI** (`excelcli`) | Coding agents | 64% fewer tokens |

**Use MCP Server when:** You're having a conversation with an AI about Excel, and you want it to discover available operations through tool schemas instead of reading skill documentation.

---

## Example Use Cases

**"Create a sales tracker with Date, Product, Quantity, Unit Price, and Total columns"**  
→ AI creates the workbook, adds headers, enters sample data, and builds formulas

**"Create a PivotTable from this data showing total sales by Product, then add a chart"**  
→ AI creates PivotTable, configures fields, and adds a linked visualization

**"Import products.csv with Power Query, load to Data Model, create a Total Revenue measure"**  
→ AI imports data, adds to Power Pivot, and creates DAX measures for analysis

**"Create a slicer for the Region field so I can filter interactively"**  
→ AI adds slicers connected to PivotTables or Tables for point-and-click filtering

**"Put this data in A1: Name, Age / Alice, 30 / Bob, 25"**  
→ AI writes data directly to cells using natural delimiters you provide

**"Show me Excel while you work"**  
→ AI makes Excel visible so you can watch changes happen in real-time

---

## Key Features

### Powered by the Real Excel Engine

ExcelMcp drives the **actual Excel application** through its official COM API — not a file parser. That means it can:

- ✅ Run live operations — refresh Power Query, recalculate, refresh PivotTables and the Data Model, evaluate DAX, run VBA and Python `=PY()`
- ✅ Edit existing workbooks safely — formulas, PivotTables, charts, macros and formatting stay intact
- ✅ Show changes live in Excel as the AI works
- ✅ Use native Excel validation and error handling

### AI-Powered Workflows

- 💬 Natural language Excel commands through AI assistants
- 🔄 Optimize Power Query M code for performance and readability
- 📊 Build complex DAX measures with AI guidance
- 📋 Automate repetitive data transformations and formatting
- 👀 **Show Excel Mode** — Watch changes live as AI works
- 🚀 **First-Run Bootstrap** — Auto-download the newest self-contained MCP runtime when the plugin is first invoked

### Optional Remote Code Formatting

- M and DAX code is preserved locally by default.
- Setting `formatMCode=true` sends M code to powerqueryformatter.com.
- Setting `formatDax=true` sends DAX to daxformatter.com.
- Remote formatting requires explicit user consent and adds network latency.

---

## Supported AI Assistants

- ✅ GitHub Copilot (VS Code, Visual Studio)
- ✅ Claude Desktop
- ✅ Cursor
- ✅ Cline (VS Code Extension)
- ✅ Windsurf
- ✅ Any MCP-compatible client

---

## Resources

- **Documentation:** [excelmcpserver.dev](https://excelmcpserver.dev/)
- **Installation Guide:** [excelmcpserver.dev/installation-mcp-server](https://excelmcpserver.dev/installation-mcp-server/)
- **GitHub Repository:** [github.com/sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel)
- **Issues:** [github.com/sbroenne/mcp-server-excel/issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **License:** MIT
