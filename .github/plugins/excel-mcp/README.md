# Excel MCP Plugin

**Model Context Protocol server for natural language Excel automation**

This plugin provides the `excel-mcp` skill and MCP server for GitHub Copilot. Use natural language to automate Power Query, DAX measures, PivotTables, Tables, Charts, VBA macros, and more through Windows Excel COM API.

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
# Register the plugin marketplace (one-time)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install the MCP plugin
copilot plugin install excel-mcp@mcp-server-excel-plugins
```

### Option 3: Manual Installation

1. Download `ExcelMcp-MCP-Server-{version}-windows.zip` from [Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Extract `mcp-excel.exe` to a permanent location
3. Add to your MCP client configuration (see [Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md))

---

## What You Can Do

**25 specialized tools with 230 operations** for comprehensive Excel automation:

### Core Operations

- **Power Query** (12 ops) — Create, update, refresh; M code formatting via powerqueryformatter.com
- **Data Model/DAX** (19 ops) — Measures, relationships, EVALUATE queries; DAX formatting via SQLBI Dax.Formatter
- **PivotTables** (30 ops) — Create, configure fields, calculated items/fields, layouts
- **Excel Tables** (27 ops) — Lifecycle, filtering, sorting, DAX-backed tables
- **Charts** (29 ops) — Create, configure, series, data labels, trendlines, axis formatting
- **Ranges** (46 ops) — Values, formulas, formatting, validation, protection, merge/unmerge
- **Worksheets** (16 ops) — Create, rename, copy, move, cross-workbook operations

### Advanced Features

- **VBA** (6 ops) — Module import/export, run procedures, version control
- **Connections** (9 ops) — OLEDB/ODBC/Text/Web connections with refresh and testing
- **Slicers** (8 ops) — Interactive filtering for PivotTables and Tables
- **Conditional Formatting** (2 ops) — Add rules (cell value, expression), clear
- **Named Ranges** (6 ops) — Create, update, delete named ranges
- **Calculation Mode** (3 ops) — Get/set mode, trigger recalculation (performance optimization)
- **Screenshot** (2 ops) — Capture ranges/sheets as PNG for visual verification
- **Window Management** (9 ops) — Show/hide Excel, arrange, position, status bar feedback

### File Operations

- **Session Management** (6 ops) — Create, open, close, list sessions
- **IRM/AIP Support** — Auto-detects protected files, opens with Excel visible for authentication

**Complete documentation:** [Full Feature Reference](https://sbroenne.github.io/mcp-server-excel/features/)

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

### Safe by Design

**100% Safe — Uses Excel's Native COM API**

Unlike third-party libraries that manipulate `.xlsx` files (risking corruption), ExcelMcp uses **Excel's official COM automation API**. This guarantees:

- ✅ Zero risk of file corruption
- ✅ Real-time changes visible in Excel
- ✅ Native Excel validation and error handling
- ✅ Full support for complex features (Power Query, DAX, VBA)

### AI-Powered Workflows

- 💬 Natural language Excel commands through AI assistants
- 🔄 Optimize Power Query M code for performance and readability
- 📊 Build complex DAX measures with AI guidance
- 📋 Automate repetitive data transformations and formatting
- 👀 **Show Excel Mode** — Watch changes live as AI works

### Automatic Code Formatting

- **M Code** — Formatted via powerqueryformatter.com (mogularGmbH, MIT License)
- **DAX** — Formatted via SQLBI Dax.Formatter library
- Adds ~100-500ms per operation for dramatically improved readability
- Graceful fallback if formatting unavailable

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

- **Documentation:** [sbroenne.github.io/mcp-server-excel](https://sbroenne.github.io/mcp-server-excel/)
- **Installation Guide:** [docs/INSTALLATION.md](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md)
- **GitHub Repository:** [github.com/sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel)
- **Issues:** [github.com/sbroenne/mcp-server-excel/issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **License:** MIT
