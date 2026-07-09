---
title: CLI Documentation
description: Full command-line interface reference for Excel automation — token-efficient Excel control for coding agents like GitHub Copilot, Cursor and Windsurf.
keywords: "Excel CLI, excelcli, command line Excel automation, coding agent Excel, Copilot CLI"
---

# CLI Documentation

{%
  include-markdown "../../src/ExcelMcp.CLI/README.md"
  start="<!--start-->"
  end="<!--end-->"
  heading-offset=1
%}

### 🤫 Quiet Mode (Agent-Friendly)

For scripting and coding agents, use `-q`/`--quiet` to suppress banner and output JSON only:

```powershell
excelcli -q session open data.xlsx
excelcli -q range get-values --session 1 --sheet Sheet1 --range A1:B2
excelcli -q session close --session 1 --save

# IRM/AIP-protected workbook
excelcli -q session open protected.xlsx --show --timeout 15
```

Banner auto-suppresses when stdout is piped or redirected.

### 🆘 Built-in Help

- `excelcli --help` – lists every command category plus the descriptions from `Program.cs`
- `excelcli <command> --help` – shows verb-specific arguments (for example `excelcli sheet --help`)
- `excelcli session --help` – displays nested verbs such as `open`, `save`, `close`, and `list`

Descriptions are kept in sync with the CLI source so the help output always reflects the latest capabilities.

### 📋 Command Categories

**Quick Reference:**

| Category | Operations | Examples |
|----------|-----------|----------|
| **File & Session** | 6 | `session create`, `session open` (IRM/AIP auto-detected), `session close`, `session list` |
| **Worksheets** | 16 | `sheet list`, `sheet create`, `sheet rename`, `sheet copy`, `sheet move`, `sheet copy-to-file` |
| **Power Query** | 12 | `powerquery list`, `powerquery create`, `powerquery refresh`, `powerquery update` |
| **Ranges** | 46 | `range get-values`, `range set-values`, `range copy`, `range find`, `range merge-cells` |
| **Conditional Formatting** | 2 | `conditionalformat add-rule`, `conditionalformat clear-rules` |
| **Excel Tables** | 27 | `table create`, `table apply-filter`, `table get-data`, `table sort`, `table add-column` |
| **Charts** | 8 | `chart create-from-range`, `chart list`, `chart delete`, `chart move`, `chart fit-to-range` |
| **Chart Config** | 21 | `chartconfig set-title`, `chartconfig add-series`, `chartconfig set-style`, `chartconfig data-labels` |
| **PivotTables** | 30 | `pivottable create-from-range`, `pivottable add-row-field`, `pivottable refresh` |
| **Slicers** | 8 | `slicer create-slicer`, `slicer list-slicers`, `slicer set-slicer-selection` |
| **Data Model** | 19 | `datamodel create-measure`, `datamodel create-relationship`, `datamodel evaluate` |
| **Connections** | 9 | `connection list`, `connection refresh`, `connection test` |
| **Named Ranges** | 6 | `namedrange create`, `namedrange read`, `namedrange write`, `namedrange update` |
| **VBA** | 6 | `vba list`, `vba import`, `vba run`, `vba update` |
| **Calculation Mode** | 3 | `calculation get-mode`, `calculation set-mode`, `calculation calculate` |
| **Python in Excel** | 2 | `pythoninexcel set-formula`, `pythoninexcel get-result` |
| **Screenshot** | 2 | `screenshot capture`, `screenshot capture-sheet` |
| **Window Management** | 9 | `window show`, `window hide`, `window arrange`, `window get-state` |

**Note:** CLI uses session commands for multi-operation workflows. See the [Feature Reference](features.md) for the full operation-by-operation documentation.

### SESSION LIFECYCLE (Open/Save/Close)

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

#### Session Lifecycle Benefits

- **Explicit control** - Know exactly when changes are persisted with `--save`
- **Batch efficiency** - Keep single Excel instance open for multiple operations (75-90% faster)
- **Flexibility** - Save and close in one command, or close without saving
- **Clean resource management** - Automatic Excel cleanup when session closes

#### Background Service & System Tray

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

### 💡 Command Reference

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

#### Typical Workflows

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

### 🛠️ Advanced Usage

#### Scripting & Automation

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

#### CI/CD Integration

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
    $session = (excelcli session open data.xlsx | Select-String "Session ID:").ToString().Split()[-1]
    excelcli powerquery create --session $session --query-name "Query1" --m-code-file queries/query1.pq
    excelcli powerquery refresh --session $session --query-name "Query1"
    excelcli session close --session $session --save
  shell: pwsh
```

### ✅ Tested Scenarios

The CLI ships with real Excel-backed integration tests that exercise the session lifecycle plus worksheet creation/listing flows through the same commands you run locally. Execute them with:

```powershell
dotnet test tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj --filter "Layer=CLI"
```

These tests open actual workbooks, issue `session open/list/close`, and call `excelcli sheet` actions to ensure the command pipeline stays healthy.

### 🔒 VBA Operations Setup

See the [Installation Guide's VBA setup section](installation-cli.md#vba-access-denied) for one-time Trust Center configuration required before using `vba` commands.

### 🚧 Troubleshooting

See the [Installation Guide's Troubleshooting section](installation-cli.md#troubleshooting) for solutions to common issues (command not found, Excel not found, VBA access denied, permission issues, IRM/AIP protected workbooks).
