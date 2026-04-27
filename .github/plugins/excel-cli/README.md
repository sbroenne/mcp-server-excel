# Excel CLI Plugin

**Command-line Excel automation for coding agents — 64% more token-efficient than MCP Server**

This plugin provides the `excel-cli` skill plus a lightweight runtime bootstrap for GitHub Copilot CLI agents. The skill guides agents to use `excelcli` commands for Power Query, DAX, PivotTables, Tables, Charts, VBA, and more — all through Windows Excel COM automation.

**Best for:** Coding agents (GitHub Copilot, Cursor, Windsurf) that need Excel automation without loading large tool schemas into context.

---

## Prerequisites

- **Windows** with Microsoft Excel 2016 or later (COM interop required)

---

## Installation

### Step 1: Register Plugin Marketplace and Install

```powershell
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

### Step 2: Install the Optional Global Shim

If you want `excelcli` on PATH for shell usage outside plugin-driven flows, install the plugin-provided shim:

```powershell
pwsh -ExecutionPolicy Bypass -File `
  "$env:USERPROFILE\.copilot\installed-plugins\mcp-server-excel-plugins\excel-cli\bin\install-global.ps1"
```

This writes `excelcli.cmd` / `excelcli.ps1` to `~/.copilot/bin` and adds that directory to your user PATH if needed.

### Step 3: First Use Bootstraps `excelcli`

The plugin now ships **wrapper/download logic** instead of a bundled executable. On first real invocation it:

1. Checks the runtime cache under `~/.copilot\plugin-runtime\mcp-server-excel\excel-cli`
2. Queries the newest GitHub Release from `sbroenne/mcp-server-excel`
3. Downloads the self-contained Windows CLI asset if needed
4. Reuses that runtime for the rest of the chat session without repeated freshness checks

You do **not** need a separate standalone install just to use the plugin.

### Step 4: Optional Standalone CLI Install

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

**17 command categories with 230 operations** for comprehensive Excel automation:

- **Power Query** (10 ops) — Create, update, refresh queries; M code management
- **Data Model/DAX** (19 ops) — Measures, relationships, EVALUATE queries
- **PivotTables** (30 ops) — Create, configure fields, calculated items/fields
- **Excel Tables** (27 ops) — Lifecycle, filtering, sorting, DAX-backed tables
- **Charts** (28 ops) — Create, configure, series, data labels, trendlines
- **Ranges** (46 ops) — Values, formulas, formatting, validation, protection
- **Worksheets** (16 ops) — Create, rename, copy, move between workbooks
- **VBA** (6 ops) — Module management and execution
- **Connections** (9 ops) — OLEDB/ODBC/Text/Web connection management
- **Conditional Formatting** (2 ops) — Add rules, clear formatting
- **Slicers** (8 ops) — Interactive filtering for PivotTables and Tables
- **Named Ranges** (6 ops) — Create, update, delete named ranges
- **Calculation Mode** (3 ops) — Get/set mode, trigger recalculation
- **Screenshot** (2 ops) — Capture ranges/sheets as PNG
- **File Operations** (6 ops) — Create, open, close, session management
- **Window Management** (9 ops) — Show/hide Excel, positioning
- **Diagnostics** (3 ops) — Health checks and troubleshooting

**Complete documentation:** [Full Feature Reference](https://sbroenne.github.io/mcp-server-excel/features/)

---

## Why CLI Over MCP Server?

| Interface | Best For | Token Efficiency |
|-----------|----------|------------------|
| **CLI** (`excelcli`) | Coding agents | **64% fewer tokens** — single tool + skill |
| **MCP Server** | Conversational AI (Claude Desktop) | 25 tool schemas loaded into context |

**Use CLI when:** Your agent needs to script Excel operations without consuming context with large tool definitions.

---

## Quick Start Example

```powershell
# Create new workbook
excelcli -q session create C:\Reports\Sales.xlsx

# Write headers
excelcli -q range set-values --session <id> --sheet Sheet1 \
  --range-address A1:C1 \
  --values '[["Date","Product","Revenue"]]'

# Write data rows
excelcli -q range set-values --session <id> --sheet Sheet1 \
  --range-address A2:C3 \
  --values '[["2024-01-15","Widget",1500],["2024-01-16","Gadget",2300]]'

# Create Excel Table
excelcli -q table create --session <id> --sheet Sheet1 \
  --table-name SalesData --range-address A1:C3

# Save and close
excelcli -q session close --session <id> --save
```

---

## Key Features

- **Zero Corruption Risk** — Uses Excel's native COM API (not file manipulation)
- **Session Management** — Open once, run many operations, close cleanly
- **Quiet Mode** (`-q`) — JSON output only, perfect for scripting
- **Built-in Help** — `excelcli --help` and `excelcli <command> --help`
- **Runtime Bootstrap** — Resolves the newest Windows CLI release once per Copilot chat session and caches it locally
- **IRM/AIP Support** — Auto-detects protected files, opens with Excel visible for sign-in

---

## Support

- **Documentation:** [sbroenne.github.io/mcp-server-excel](https://sbroenne.github.io/mcp-server-excel/)
- **Issues:** [github.com/sbroenne/mcp-server-excel/issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **License:** MIT
