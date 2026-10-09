# Excel CLI Plugin

**Command-line Excel automation for coding agents — 64% more token-efficient than MCP Server**

This plugin provides an npx-first `excelcli` launcher, a small `excel-cli` skill
that helps agents discover that launcher for ordinary workbook requests, and the optional
`excel-cli-report-formatting` skill for requested report presentation.
Ordinary Excel automation uses native CLI help; general workflows and recovery
remain in the [documentation](https://excelmcpserver.dev/reference/).

**Best for:** Coding agents (GitHub Copilot, Cursor, Windsurf) that need Excel automation without loading large tool schemas into context.

---

## Prerequisites

- **Windows** with Microsoft Excel 2016 or later (COM interop required)
- **Node.js 18 or later** with `npx`

---

## Installation

### Step 1: Install the Plugin

Install from [Awesome Copilot](https://github.com/github/awesome-copilot), the
default marketplace in current Copilot clients:

```powershell
copilot plugin install excel-cli@awesome-copilot
```

Alternatively, install from our direct marketplace:

```powershell
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

Choose one marketplace for this plugin; do not install both copies.

### Step 2: Run `excelcli` through npm

The plugin does not bundle `excelcli.exe`. Its wrapper runs:

```powershell
npx -y @sbroenne/excelcli@latest --help
```

Node.js and npx are required. The plugin's `bin\start-cli.ps1` wrapper preserves
quoted JSON arguments when invoked from Windows PowerShell. No global helper,
PATH change, or separate runtime installation is required.

You do **not** need a separate standalone install just to use the plugin.

### Optional Standalone CLI Install

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

**31 feature command categories with 393 operations** for comprehensive Excel automation:

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
| **MCP Server** | Conversational AI (Claude Desktop) | 60 tool schemas loaded into context |

**Use CLI when:** Your agent needs to script Excel operations without consuming context with large tool definitions.

---

## Quick Start Example

The examples below use `excelcli` for readability. Plugin installation does not
put that command on PATH: replace it with `npx -y @sbroenne/excelcli@latest`
unless you installed a standalone CLI. For quoted JSON arguments in Windows
PowerShell, use the plugin's `bin\start-cli.ps1` wrapper as the command instead:

Use your client's installed plugin directory rather than assuming a
marketplace-specific path. Replace the example directory below:

```powershell
& "C:\Path\To\Installed\excel-cli\bin\start-cli.ps1" --help
```

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
- **Built-in Help** — `npx -y @sbroenne/excelcli@latest --help` and `npx -y @sbroenne/excelcli@latest <command> --help`
- **npm Launch** — Uses the npm `latest` tag; npm manages package resolution and caching subject to its cache policy
- **IRM/AIP Support** — Auto-detects protected files, opens with Excel visible for sign-in

---

## Support

- **Documentation:** [excelmcpserver.dev](https://excelmcpserver.dev/)
- **Issues:** [github.com/sbroenne/mcp-server-excel/issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **License:** MIT
