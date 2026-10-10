# ExcelMcp.CLI - Command-Line Interface for Excel Automation

[![GitHub Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![GitHub Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=Downloads)](https://github.com/sbroenne/mcp-server-excel/releases)
[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.CLI.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Command-line interface for Excel automation — preferred by coding agents.**

> **Primary distribution: npm or standalone executable** — Run `npx -y @sbroenne/excelcli@latest --help` or download `excelcli.exe` from the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest). No .NET runtime required.
> **Secondary distribution: NuGet .NET tool** — `dotnet tool install --global Sbroenne.ExcelMcp.CLI` (requires .NET 10 runtime).

The CLI provides 31 feature command categories with 393 operations matching the MCP Server, plus `session`, `service`, and `batch` commands — the same capabilities without loading 31 tool schemas into context.

| Interface | Best For | Why |
|-----------|----------|-----|
| **CLI** (`excelcli`) | Coding agents (Copilot, Cursor, Windsurf) | **64% fewer tokens** - single tool, no large schemas |
| **MCP Server** | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent connection |

Also perfect for RPA workflows, CI/CD pipelines, batch processing, and automated testing.

➡️ **[Learn more and see examples](https://excelmcpserver.dev/)**

---

## 🚀 Quick Start

### Primary Installation: npm

```powershell
npx -y @sbroenne/excelcli@latest --help
# Or install the command on your PATH:
npm install --global @sbroenne/excelcli@latest
excelcli --version
```

Requires Node.js 18+. Keep optional dependencies enabled so npm installs the
matching Windows runtime. ARM64 Node.js selects the native ARM64 package;
x64 Node.js selects the x64 package, which runs through emulation on ARM64
Windows. A missing matching runtime fails with reinstall guidance rather than
falling back to another architecture. CLI arguments follow the package name
when using `npx`, for example `npx -y @sbroenne/excelcli@latest -q session list`.
`@latest` uses normal npm caching and does not replace a running CLI service.
Finish and explicitly save/close workbook sessions before stopping the old
service for an update.

### Primary Installation: Standalone Executable

1. Download **`ExcelMcp-CLI-{version}-windows.zip`** from the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Extract `excelcli.exe` to a permanent location (e.g., `C:\Tools\ExcelMcp\`) and add the directory to your PATH
3. Verify: `excelcli --version` and `excelcli --help`

### Secondary Installation: .NET Global Tool

```powershell
# Requires .NET 10 Runtime or SDK
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

📖 **[Full Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION-CLI.md)** - PATH setup, GitHub Copilot plugin, updating, uninstalling, and troubleshooting

📚 **CLI usage guide:** See the session workflow, troubleshooting, advanced usage, and CI/CD examples below.

> 🔁 **Session Workflow:** Always start with `excelcli session open <file>` (captures the session id), pass `--session <id>` to other commands, then `excelcli session close --session <id> --save` when finished. Add `--show` when Excel must stay visible for IRM/AIP sign-in or other authentication prompts.

---

## 📋 What You Can Do

ExcelMcp.CLI provides **393 operations** across 31 feature command categories including Power Query, Python in Excel, Data Model/DAX, What-If Analysis, PivotTables, Excel Tables, Charts, Drawings, VBA, Ranges, Worksheets, Workbooks, QueryTables, XML Maps, Connections, and Window Management.

Drives the **actual Excel application** via COM — not a file-format parser — so live operations (Power Query refresh, recalculation, DAX evaluation, VBA execution) run for real and existing workbooks stay intact.

📚 **[Feature Overview →](https://github.com/sbroenne/mcp-server-excel/blob/main/FEATURES.md)** - Capabilities grouped by goal. Use `excelcli --help` and `excelcli <command> --help` for current actions and inputs.

---

## ⚙️ System Requirements

- **Windows OS** (Windows 10/11 or Server 2016+) + **Microsoft Excel 2016 or later** — COM interop is Windows-specific and requires Excel to be installed
- **Node.js 18+** only if using npm; Windows x64 and ARM64 are supported, with the runtime selected by Node.js architecture
- **.NET 10 Runtime** only if using the NuGet .NET tool install path (not required for npm or the standalone exe)

📖 **[Full System Requirements & Optional Components](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION-CLI.md)** - including DAX/MSOLAP prerequisites

---

## 📖 Complete Documentation

- **[npm Package](https://www.npmjs.com/package/@sbroenne/excelcli)** - Self-contained CLI through npm (primary)
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
# Solution: In Excel, enable File → Options → Trust Center → Trust Center Settings
#           → Macro Settings → "Trust access to the VBA project object model"
```

### Permission Issues

```powershell
# Run PowerShell/CMD as Administrator if you encounter permission errors
# excelcli.exe is a standalone exe - no installation needed
```

### IRM / AIP Protected Workbooks

```powershell
# Validate an ordinary workbook through a temporary read-only Excel open
excelcli -q session test "D:\Docs\Workbook.xlsx" --timeout-seconds 120

# Inspect deterministic protection requirements before an interactive open
excelcli -q session test "D:\Docs\Protected.xlsx"

# Keep Excel visible so authentication or policy prompts can surface
excelcli session open "D:\Docs\Protected.xlsx" --show --timeout-seconds 120
```

`session test` reports `canOpen`, `isIrmProtected`, `willOpenReadOnly`, and
`requiresVisibleSession` using the same result model as MCP `file_read` action `test`. Protected
files report `canOpen:false` until interactive Excel authentication occurs. Use
`--show` whenever hidden automation would block on a sign-in, consent, or
information-protection prompt. Ordinary files are opened read-only in a temporary
Excel session and closed without saving.

Protection detection does not force read-only access. After visible authentication,
Excel decides the signed-in user's editing rights. `willOpenReadOnly:false` is not
a guarantee of edit access; inspect `readOnly` with `excelcli workbook get-info`
(MCP: `workbook_read` action `get-info`) before editing. Failed saves return an
error and leave the session open with any unsaved changes available for inspection.

### Daemon Status and Session Discovery

`excelcli -q service status` reports `daemonState` as `stopped`, `starting`,
`running`, or `unresponsive`. A stopped daemon is a successful status result with
`running:false`; a transport timeout is an error with `running:true` and
`daemonState:"unresponsive"`.

`excelcli -q session list` returns an empty `sessions` array only when the daemon
is confirmed stopped or a responsive daemon confirms it has no sessions.
Transport failures exit nonzero without a `sessions` property. Status and list
allow up to 10 seconds for daemon transport readiness, while daemon startup
allows up to 30 seconds.

---

## 🛠️ Advanced Usage

### Scripting & Automation

```powershell
# PowerShell script example
$files = Get-ChildItem *.xlsx
foreach ($file in $files) {
    $sessionId = (excelcli -q session open $file.FullName | ConvertFrom-Json).sessionId
    excelcli -q powerquery refresh --session $sessionId --query-name "Sales Data"
    excelcli -q datamodel refresh --session $sessionId
    excelcli -q session close --session $sessionId --save
}
```

Each command prints a JSON result and exits with code 0 on success. It exits
with code 1 when the command fails or its result reports `success: false` (for
example, a `powerquery refresh-all` where one query failed), so scripts can
check `$LASTEXITCODE`. In `batch` output, such a line has `success: false`, and
`--stop-on-error` stops there.

### Incremental batch input

`excelcli -q batch --stream` keeps one client running for dependent commands.
Send one complete JSON object per line on stdin; each result is flushed as one
NDJSON line before the client waits for the next command. The caller can inspect
that result and choose the next command while keeping stdin open. Closing stdin
ends the client; it does not save or close open workbook sessions.

Use the same `group.action`, camelCase argument names, session selection, and
`--stop-on-error` policy as ordinary batch. `--stream` accepts stdin only
(`--input -` is also allowed); JSON arrays and multiline objects are not valid
stream entries. Blank lines are ignored. Failed entries produce indexed error
results, and the final exit code is nonzero if any entry failed. If the service
cannot be reached, that entry's result reports the communication error and the
next command retries the connection. Cancellation
ends the invocation. Inspect any partial workbook changes before continuing.

Without `--stream`, batch still reads the complete input before executing.
An agent needs process tools that can write further input and read output from
a retained process to use incremental mode. Discover unfamiliar command syntax
with normal `--help`; streaming input does not add a help action.

### CI/CD Integration

Excel COM requires a self-hosted Windows runner with desktop Excel installed; GitHub-hosted runners do not include Excel.

```yaml
# GitHub Actions example
jobs:
  process-excel:
    runs-on: [self-hosted, Windows, excel]
    steps:
      - name: Download ExcelMcp CLI
        shell: pwsh
        run: |
          $version = (Invoke-RestMethod "https://api.github.com/repos/sbroenne/mcp-server-excel/releases/latest").tag_name.TrimStart('v')
          Invoke-WebRequest "https://github.com/sbroenne/mcp-server-excel/releases/download/v$version/ExcelMcp-CLI-$version-windows.zip" -OutFile cli.zip
          Expand-Archive cli.zip -DestinationPath C:\Tools\ExcelMcp
          "C:\Tools\ExcelMcp" >> $env:GITHUB_PATH

      - name: Process Excel Files
        shell: pwsh
        run: |
          $sessionId = (excelcli -q session open data.xlsx | ConvertFrom-Json).sessionId
          excelcli -q powerquery create --session $sessionId --query-name "Query1" --m-code-file queries\query1.pq
          excelcli -q powerquery refresh --session $sessionId --query-name "Query1"
          excelcli -q session close --session $sessionId --save
```


## ✅ Tested Scenarios

The CLI ships with real Excel-backed integration tests that exercise the session lifecycle plus worksheet creation/listing flows through the same commands you run locally. Execute them with:

```powershell
dotnet test tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj --filter "Layer=CLI"
```

These tests open actual workbooks, issue `session open/list/close`, and call `excelcli sheet` actions to ensure the command pipeline stays healthy.

---

## 🤝 Related Tools

- **[MCP Server](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.McpServer/README.md)** - For conversational AI (Claude Desktop, VS Code Chat) — distributed as `mcp-excel.exe`
- **[VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)** - One-click Excel automation in VS Code
- **Issues & Discussions**: [GitHub](https://github.com/sbroenne/mcp-server-excel)
- **Full docs**: [excelmcpserver.dev](https://excelmcpserver.dev/)

---

## 📄 License

MIT License - see [LICENSE](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE) for details.

---

**Built with ❤️ for Excel developers and automation engineers**
