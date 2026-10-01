# ExcelMcp CLI

Run the self-contained Excel automation CLI through npm:

```powershell
npx -y @sbroenne/excelcli@latest --help
npx -y @sbroenne/excelcli@latest -q session open "C:\Data\Book.xlsx"
```

On Mac, use your actual absolute native workbook path:

```bash
npx -y @sbroenne/excelcli@latest -q session open "/absolute/path/Book.xlsx"
```

For repeated use, install the `excelcli` command on your PATH:

```powershell
npm install --global @sbroenne/excelcli@latest
excelcli --version
```

Requires Node.js 18 or later and either Windows x64/ARM64 with
Microsoft Excel 2016+, or Apple Silicon macOS with Excel for Mac 16.112+.
No separate .NET runtime is needed. Keep
optional dependencies enabled so npm installs the matching native runtime.
ARM64 Node.js selects the native ARM64 runtime; x64 Node.js selects x64,
including on ARM64 Windows. A missing matching package is an error, not a
fallback to another architecture.

The launcher forwards arguments, standard input/output, and exit codes to the
existing CLI. Excel operations and session management are unchanged.

`@latest` selects the current npm release using normal npm caching. It does not
replace an already running background service. Finish and explicitly save/close
workbook sessions before stopping that service to use a new version.
Network access is needed for package downloads and update checks.

## macOS: experimental beta

Apple Silicon support is a limited beta, not Windows feature parity. Power
Query, VBA, Data Model/DAX/OLAP, Tables, PivotTables, charts, slicers,
connections, QueryTables, XML Maps, screenshots, advanced visual formatting,
and Python result reads are unavailable. Test on copies of important workbooks.
See [macOS beta limitations](https://github.com/sbroenne/mcp-server-excel/blob/main/specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta)
and the [per-action inventory](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/MACOS-ACTION-INVENTORY.md).
Windows retains the complete backend.

[Documentation](https://excelmcpserver.dev/installation-cli/) |
[Source](https://github.com/sbroenne/mcp-server-excel) |
[Issues](https://github.com/sbroenne/mcp-server-excel/issues)
