# ExcelMcp.CLI - Command-Line Interface for Excel Automation

[![GitHub Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![GitHub Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=Downloads)](https://github.com/sbroenne/mcp-server-excel/releases)
[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.CLI.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Command-line interface for Excel automation — preferred by coding agents.**

> **Primary distribution: Standalone executable** — Download `excelcli.exe` from the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest). No .NET runtime required.
> **Secondary distribution: NuGet .NET tool** — `dotnet tool install --global Sbroenne.ExcelMcp.CLI` (requires .NET 10 runtime).

The CLI provides 17 command categories with 230 operations matching the MCP Server. Uses **64% fewer tokens** than MCP Server because it wraps all operations in a single tool with skill-based guidance instead of loading 25 tool schemas into context.

| Interface | Best For | Why |
|-----------|----------|-----|
| **CLI** (`excelcli`) | Coding agents (Copilot, Cursor, Windsurf) | **64% fewer tokens** - single tool, no large schemas |
| **MCP Server** | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent connection |

Also perfect for RPA workflows, CI/CD pipelines, batch processing, and automated testing.

➡️ **[Learn more and see examples](https://sbroenne.github.io/mcp-server-excel/)**

---

## 🚀 Quick Start

### Primary Installation: Standalone Executable

1. Download **`ExcelMcp-CLI-{version}-windows.zip`** from the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Extract `excelcli.exe` to a permanent location (e.g., `C:\Tools\ExcelMcp\`)
3. Add the directory to your PATH

```powershell
# Verify installation
excelcli --version

# Get help
excelcli --help
```

> 🔁 **Session Workflow:** Always start with `excelcli session open <file>` (captures the session id), pass `--session <id>` to other commands, then `excelcli session close --session <id> --save` when finished. Add `--show` when Excel must stay visible for IRM/AIP sign-in or other authentication prompts.

### Secondary Installation: .NET Global Tool

```powershell
# Requires .NET 10 Runtime or SDK
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

### Check for Updates

```powershell
# Check if newer version is available
excelcli --version
# If an update is available, download the latest release from:
# https://github.com/sbroenne/mcp-server-excel/releases/latest

# Or update via NuGet (if installed that way):
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

### Uninstall

```powershell
# Standalone exe: delete the file
Remove-Item "C:\Tools\ExcelMcp\excelcli.exe" -Force

# NuGet (if installed via dotnet tool):
dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI
```

## 🤫 Quiet Mode (Agent-Friendly)

For scripting and coding agents, use `-q`/`--quiet` to suppress banner and output JSON only:

```powershell
excelcli -q session open data.xlsx
excelcli -q range get-values --session 1 --sheet Sheet1 --range A1:B2
excelcli -q session close --session 1 --save

# IRM/AIP-protected workbook
excelcli -q session open protected.xlsx --show --timeout 15
```

Banner auto-suppresses when stdout is piped or redirected.

## 🆘 Built-in Help

- `excelcli --help` – lists every command category plus the new descriptions from `Program.cs`
- `excelcli <command> --help` – shows verb-specific arguments (for example `excelcli sheet --help`)
- `excelcli session --help` – displays nested verbs such as `open`, `save`, `close`, and `list`

Descriptions are kept in sync with the CLI source so the help output always reflects the latest capabilities.

---

## ✨ Key Features

### 🔧 Excel Development Automation
- **Power Query Management** - Export, import, update, and version control M code
- **VBA Development** - Manage VBA modules and run procedures in `.xlsm` workbooks
- **Data Model & DAX** - Create measures, manage relationships, Power Pivot operations
- **PivotTable Automation** - Create, configure, and manage PivotTables programmatically
- **Conditional Formatting** - Add rules (cell value, expression-based), clear formatting

### 📊 Data Operations
- **Worksheet Management** - Create, rename, copy, delete sheets with tab colors and visibility
- **Range Operations** - Read/write values, formulas, formatting, validation
- **Excel Tables** - Lifecycle management, filtering, sorting, structured references
- **Connection Management** - OLEDB, ODBC, Text, Web connections with testing

### 🛡️ Production Ready
- **Zero Corruption Risk** - Uses Excel's native COM API (not file manipulation)
- **Error Handling** - Comprehensive validation and helpful error messages
- **CI/CD Integration** - Perfect for automated workflows and testing
- **Windows Native** - Optimized for Windows Excel automation

---

## 📋 Command Categories

ExcelMcp.CLI provides **230 operations** across 17 command categories:

📚 **[Complete Feature Reference →](../../FEATURES.md)** - Full documentation with all operations

**Quick Reference:**

| Category | Operations | Examples |
|----------|-----------|----------|
| **File & Session** | 6 | `session create`, `session open` (IRM/AIP auto-detected), `session close`, `session list` |
| **Worksheets** | 16 | `sheet list`, `sheet create`, `sheet rename`, `sheet copy`, `sheet move`, `sheet copy-to-file` |
| **Power Query** | 10 | `powerquery list`, `powerquery create`, `powerquery refresh`, `powerquery update` |
| **Ranges** | 46 | `range get-values`, `range set-values`, `range copy`, `range find`, `range merge-cells` |
| **Conditional Formatting** | 2 | `conditionalformat add-rule`, `conditionalformat clear-rules` |
| **Excel Tables** | 27 | `table create`, `table apply-filter`, `table get-data`, `table sort`, `table add-column` |
| **Charts** | 14 | `chart create-from-range`, `chart list`, `chart delete`, `chart move`, `chart fit-to-range` |
| **Chart Config** | 14 | `chartconfig set-title`, `chartconfig add-series`, `chartconfig set-style`, `chartconfig data-labels` |
| **PivotTables** | 30 | `pivottable create-from-range`, `pivottable add-row-field`, `pivottable refresh` |
| **Slicers** | 8 | `slicer create-slicer`, `slicer list-slicers`, `slicer set-slicer-selection` |
| **Data Model** | 19 | `datamodel create-measure`, `datamodel create-relationship`, `datamodel evaluate` |
| **Connections** | 9 | `connection list`, `connection refresh`, `connection test` |
| **Named Ranges** | 6 | `namedrange create`, `namedrange read`, `namedrange write`, `namedrange update` |
| **VBA** | 6 | `vba list`, `vba import`, `vba run`, `vba update` |
| **Calculation Mode** | 3 | `calculation get-mode`, `calculation set-mode`, `calculation calculate` |
| **Screenshot** | 2 | `screenshot capture`, `screenshot capture-sheet` |

**Note:** CLI uses session commands for multi-operation workflows.

---

## SESSION LIFECYCLE (Open/Save/Close)

The CLI uses an explicit session-based workflow where you open a file, perform operations, and optionally save before closing:

```powershell
# 1. Open a session
excelcli session open data.xlsx
# Output: Session ID: 550e8400-e29b-41d4-a716-446655440000

# IRM/AIP or auth prompt workflow
excelcli session open protected.xlsx --show

# 2. List active sessions anytime
excelcli session list

# 3. Use the session ID with any commands
excelcli sheet create --session 550e8400-e29b-41d4-a716-446655440000 --sheet "NewSheet"
excelcli powerquery list --session 550e8400-e29b-41d4-a716-446655440000

# 4. Close and save changes
excelcli session close --session 550e8400-e29b-41d4-a716-446655440000 --save

# OR: Close and discard changes (no --save flag)
excelcli session close --session 550e8400-e29b-41d4-a716-446655440000
```

### Session Lifecycle Benefits

- **Explicit control** - Know exactly when changes are persisted with `--save`
- **Batch efficiency** - Keep single Excel instance open for multiple operations (75-90% faster)
- **Flexibility** - Save and close in one command, or close without saving
- **Clean resource management** - Automatic Excel cleanup when session closes

### Background Service & System Tray

When you run your first CLI command, the **ExcelMCP Service** starts automatically in the background. The service:

- **Manages Excel COM** - Keeps Excel instance alive between commands (no restart overhead)
- **Shows system tray icon** - Look for the Excel icon in your Windows taskbar notification area
- **Tracks sessions** - Right-click the tray icon to see active sessions and close them
- **Shows session origin** - Sessions are labeled [CLI] or [MCP] showing which client created them
- **Auto-updates** - Notifies you when a new version is available and allows one-click updates

**Tray Icon Features:**
- 📋 **View sessions** - Double-click to see active session count
- 💾 **Close sessions** - Right-click → Sessions → select file → "Close Session..." (prompts to save with Cancel option)
- 🔄 **Update CLI** - When updates are available, click "Update to X.X.X" to update automatically
- ℹ️ **About** - Right-click → "About..." to see version info and helpful links
- 🛑 **Stop Service** - Right-click → "Stop Service" (prompts to save active sessions with Cancel option)

The service auto-stops after 10 minutes of inactivity (no active sessions).

---

## 💡 Command Reference

**Use `excelcli <command> --help` for complete parameter documentation.** The CLI help is always in sync with the code.

```powershell
excelcli --help              # List all commands
excelcli session --help      # Session lifecycle (open, close, save, list)
excelcli powerquery --help   # Power Query operations
excelcli range --help        # Cell/range operations
excelcli table --help        # Excel Table operations
excelcli pivottable --help   # PivotTable operations
excelcli datamodel --help    # Data Model & DAX
excelcli vba --help          # VBA module management
```

### Typical Workflows

**Session-based automation (recommended):**
```powershell
excelcli -q session open report.xlsx           # Returns session ID
excelcli -q sheet create --session 1 --sheet "Summary"
excelcli -q range set-values --session 1 --sheet Summary --range A1 --values '[["Hello"]]'
excelcli -q session close --session 1 --save   # Persist changes
```

**Visible Excel for IRM/auth workflows:**
```powershell
excelcli -q session open "D:\Docs\Protected.xlsx" --show --timeout 120
excelcli -q session list                        # session shows isExcelVisible=true
excelcli -q session close --session <id>
```

**Power Query ETL:**
```powershell
excelcli powerquery create --session 1 --query-name "CleanData" --m-code-file transform.pq
excelcli powerquery refresh --session 1 --query-name "CleanData"
```

**PivotTable from Data Model:**
```powershell
excelcli pivottable create-from-datamodel --session 1 --table Sales --dest-sheet Analysis --dest-cell A1 --pivot-table SalesPivot
excelcli pivottable add-row-field --session 1 --pivot-table SalesPivot --field Region
excelcli pivottable add-value-field --session 1 --pivot-table SalesPivot --field Amount --function Sum
```

**VBA automation:**
```powershell
excelcli vba import --session 1 --module-name "Helpers" --vba-code-file helpers.vba
excelcli vba run --session 1 --procedure-name "Helpers.ProcessData"
```

---

## ⚙️ System Requirements

| Requirement | Details | Why Required |
|-------------|---------|--------------|
| **Windows OS** | Windows 10/11 or Server 2016+ | COM interop is Windows-specific |
| **Microsoft Excel** | Excel 2016 or later | CLI controls actual Excel application |
| **.NET 10 Runtime** | [Download](https://dotnet.microsoft.com/download/dotnet/10.0) | Required to run .NET global tools |

> **Note:** ExcelMcp.CLI controls the actual Excel application via COM interop, not just file formats. This provides access to Power Query, VBA runtime, formula engine, and all Excel features, but requires Excel to be installed.

---

## 🔒 VBA Operations Setup (One-Time)

VBA commands require **"Trust access to the VBA project object model"** to be enabled manually in Excel:

1. Open Excel
2. Go to **File → Options → Trust Center**
3. Click **"Trust Center Settings"**
4. Select **"Macro Settings"**
5. Check **"✓ Trust access to the VBA project object model"**
6. Click **OK** twice

This is a security setting that must be enabled manually. ExcelMcp.CLI does not provide a `setup-vba-trust` or `check-vba-trust` command and never modifies Trust Center settings automatically.

Current VBA support is procedural and module-focused:
- `vba list` and `vba view` inspect existing VBA components and procedures
- `vba import` creates a new standard module from inline code or `--vba-code-file`
- `vba update`, `vba delete`, and `vba run` work against existing component/procedure names

For macro-enabled workbooks, use `.xlsm` extension:

```powershell
excelcli session create macros.xlsm
# Returns session ID (e.g., 1)
excelcli vba import --session 1 --module-name MyModule --vba-code-file code.vba
excelcli session close --session 1 --save
```

---

## 📖 Complete Documentation

- **[GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest)** - Download latest standalone exe (primary)
- **[NuGet Package](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)** - .NET Global Tool (secondary)
- **[GitHub Repository](https://github.com/sbroenne/mcp-server-excel)** - Source code and issues
- **[Release Notes](https://github.com/sbroenne/mcp-server-excel/releases)** - Latest updates

---

## 🚧 Troubleshooting

### Command Not Found After Installation

```powershell
# Check excelcli.exe location
where.exe excelcli

# If not found, ensure the directory containing excelcli.exe is in your PATH
# The default location after extraction might be: C:\Tools\ExcelMcp\
```

### Excel Not Found

```powershell
# Error: "Microsoft Excel is not installed"
# Solution: Install Microsoft Excel (any version 2016+)
```

### VBA Access Denied

```powershell
# Error: "Programmatic access to Visual Basic Project is not trusted"
# Solution: Enable VBA trust (see VBA Operations Setup above)
```

### Permission Issues

```powershell
# Run PowerShell/CMD as Administrator if you encounter permission errors
# excelcli.exe is a standalone exe - no installation needed
```

### IRM / AIP Protected Workbooks

```powershell
# Keep Excel visible so authentication or policy prompts can surface
excelcli session open "D:\Docs\Protected.xlsx" --show --timeout 120
```

Use `--show` whenever hidden automation would block on a sign-in, consent, or information-protection prompt.

---

## 🛠️ Advanced Usage

### Scripting & Automation

```powershell
# PowerShell script example
$files = Get-ChildItem *.xlsx
foreach ($file in $files) {
    $session = excelcli session open $file.Name | Select-String "Session ID: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    excelcli powerquery refresh --session $session --query "Sales Data"
    excelcli datamodel refresh --session $session
    excelcli session close $session --save
}
```

### CI/CD Integration

```yaml
# GitHub Actions example
- name: Download ExcelMcp CLI
  run: |
    $version = (Invoke-RestMethod "https://api.github.com/repos/sbroenne/mcp-server-excel/releases/latest").tag_name.TrimStart('v')
    Invoke-WebRequest "https://github.com/sbroenne/mcp-server-excel/releases/download/v$version/ExcelMcp-CLI-$version-windows.zip" -OutFile cli.zip
    Expand-Archive cli.zip -DestinationPath C:\Tools\ExcelMcp
    echo "C:\Tools\ExcelMcp" >> $env:GITHUB_PATH
  shell: pwsh

- name: Process Excel Files
  run: |
    SESSION=$(excelcli session open data.xlsx | grep "Session ID:" | cut -d' ' -f3)
    excelcli powerquery create --session $SESSION --query-name "Query1" --m-code-file queries/query1.pq
    excelcli powerquery refresh --session $SESSION --query-name "Query1"
    excelcli session close $SESSION --save
```


## ✅ Tested Scenarios

The CLI ships with real Excel-backed integration tests that exercise the session lifecycle plus worksheet creation/listing flows through the same commands you run locally. Execute them with:

```powershell
dotnet test tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj --filter "Layer=CLI"
```

These tests open actual workbooks, issue `session open/list/close`, and call `excelcli sheet` actions to ensure the command pipeline stays healthy.

---

## 🤝 Related Tools

- **[ExcelMcp MCP Server](https://github.com/sbroenne/mcp-server-excel/releases/latest)** - MCP server for AI assistant integration (`mcp-excel.exe`)
- **[Excel MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)** - One-click Excel automation in VS Code


---

## 📄 License

MIT License - see [LICENSE](../../LICENSE) for details.

---

## 🙋 Support

- **Issues**: [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **Discussions**: [GitHub Discussions](https://github.com/sbroenne/mcp-server-excel/discussions)
- **Documentation**: [Complete Docs](../../docs/)

---

**Built with ❤️ for Excel developers and automation engineers**
