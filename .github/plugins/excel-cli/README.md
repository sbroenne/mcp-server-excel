# Excel CLI Plugin

**Command-line Excel automation for coding agents — 64% more token-efficient than MCP Server**

This plugin provides the `excel-cli` skill for GitHub Copilot CLI agents. The
skill guides agents to invoke `npx -y @sbroenne/excelcli` for Power Query, DAX,
PivotTables, Tables, Charts, VBA, and more through the matching platform backend.

**Best for:** Coding agents (GitHub Copilot, Cursor, Windsurf) that need Excel automation without loading large tool schemas into context.

---

## Prerequisites

- **Windows x64** with Microsoft Excel 2016 or later, or **Apple Silicon macOS**
  with Excel for Mac 16.112 or later
- **Node.js 18 or later** with optional dependencies enabled

---

## Installation

### Step 1: Register Plugin Marketplace and Install

```powershell
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

### Step 2: Run the CLI

```powershell
npx -y @sbroenne/excelcli --help
```

For a persistent command on PATH, use
`npm install --global @sbroenne/excelcli`. Windows installs
`@sbroenne/excelcli-win32-x64`; Apple Silicon macOS installs
`@sbroenne/excelcli-darwin-arm64`. Unsupported platforms fail closed.

### Step 3: Optional Standalone CLI Install

If you still prefer a fully separate non-plugin install, you can use the normal release channels:

**Option A: Standalone Executable**
1. Download `ExcelMcp-CLI-{version}-windows.zip` from [Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Extract `excelcli.exe` to a permanent folder (for example `C:\Tools\ExcelMcp\`)
3. Add that folder to your PATH

**Option B: .NET Global Tool**
```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
# Requires .NET 10 Runtime
```

---

## What You Can Do

**31 feature command categories with 326 operations** for comprehensive Excel automation:

- **Power Query** (12 ops) — Create, update, refresh queries; M code management
- **Data Model/DAX** (20 ops) — Measures, relationships, source metadata, EVALUATE queries
- **PivotTables** (35 ops) — Fields, grouping, cache options, drill-through
- **Excel Tables** (27 ops) — Lifecycle, filtering, sorting, DAX-backed tables
- **Charts and Chart Config** (33 ops) — Combo series, plotting, placement, formatting
- **Ranges** (51 ops) — Values, formulas, hyperlinks, threaded comments, formatting
- **Worksheets** (33 ops) — Lifecycle, outlines, protection, notes, images, shapes
- **Workbooks** (15 ops) — Metadata, properties, Save As/copy, PDF/XPS, external links
- **VBA** (6 ops) — Module management and execution
- **Connections** (11 ops) — OLEDB/ODBC management and refresh control
- **QueryTables** (9 ops) — Text/CSV and legacy HTML imports
- **Drawing Objects** (14 ops) — Shapes, Forms controls, and sparklines
- **What-If Analysis** (8 ops) — Goal Seek, scenarios, summaries, data tables
- **XML Maps** (6 ops) — Schemas, XPath mapping, secure import/export
- **Conditional Formatting** (4 ops) — Add, list, delete, and clear rules
- **Slicers** (8 ops) — Interactive filtering for PivotTables and Tables
- **Named Ranges** (6 ops) — Create, update, delete named ranges
- **Calculation Mode** (3 ops) — Get/set mode, trigger recalculation
- **Python in Excel** (2 ops) — Set/get Python formulas and results
- **Screenshot** (2 ops) — Capture ranges/sheets as PNG
- **File Operations** (5 ops) — Create, open, close, list, and test files
- **Window Management** (15 ops) — Show/hide, panes, zoom, display options, positioning

**Complete documentation:** [Full Feature Reference](https://excelmcpserver.dev/features/)

---

## Why CLI Over MCP Server?

| Interface | Best For | Token Efficiency |
|-----------|----------|------------------|
| **CLI** (`excelcli`) | Coding agents | **64% fewer tokens** — single tool + skill |
| **MCP Server** | Conversational AI (Claude Desktop) | 31 tool schemas loaded into context |

**Use CLI when:** Your agent needs to script Excel operations without consuming context with large tool definitions.

---

## Quick Start Example

```powershell
# Create new workbook
excelcli -q session create C:\Reports\Sales.xlsx

# Write headers
excelcli -q range set-values --session <id> --sheet Sheet1 `
  --range A1:C1 `
  --values '[["Date","Product","Revenue"]]'

# Write data rows
excelcli -q range set-values --session <id> --sheet Sheet1 `
  --range A2:C3 `
  --values '[["2024-01-15","Widget",1500],["2024-01-16","Gadget",2300]]'

# Create Excel Table
excelcli -q table create --session <id> --sheet Sheet1 `
  --table-name SalesData --range A1:C3

# Save and close
excelcli -q session close --session <id> --save
```

---

## Key Features

- **Real Excel Engine** — Drives the actual Excel application via COM, so live operations run for real and existing workbooks stay intact
- **Session Management** — Open once, run many operations, close cleanly
- **Quiet Mode** (`-q`) — JSON output only, perfect for scripting
- **Built-in Help** — `excelcli --help` and `excelcli <command> --help`
- **Runtime Bootstrap** — Uses the persistent plugin cache and resolves release freshness once per Copilot chat session
- **IRM/AIP Support** — Auto-detects protected files, opens with Excel visible for sign-in

---

## Support

- **Documentation:** [excelmcpserver.dev](https://excelmcpserver.dev/)
- **Issues:** [github.com/sbroenne/mcp-server-excel/issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **License:** MIT
