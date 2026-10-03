# ExcelMcp MCP Server

Run the self-contained ExcelMcp server through npm:

```powershell
npx -y @sbroenne/mcp-server-excel@latest
```

The package supports Windows x64/ARM64 with Microsoft
Excel 2016 or later and Apple Silicon macOS with Excel for Mac 16.112 or later.
It does not require the .NET SDK or a separately installed .NET runtime. Keep
optional dependencies enabled so npm installs the matching native runtime
package.
ARM64 Node.js selects the native ARM64 runtime; x64 Node.js selects x64,
including on ARM64 Windows. A missing
matching package is an error, not a fallback to another architecture.

The Node.js entry point only launches the packaged .NET server. MCP tools and
Excel automation continue to run in the existing ExcelMcp implementation.

`@latest` selects the current npm release at startup using normal npm caching.
Network access is needed for downloads and update checks. Restart your MCP
server after safely finishing workbook work to run an updated version.

## macOS: experimental beta

Apple Silicon support is a limited beta, not Windows feature parity. Power
Query, VBA, Data Model/DAX/OLAP, Tables, PivotTables, charts, slicers,
connections, QueryTables, XML Maps, screenshots, advanced visual formatting,
and Python result reads are unavailable. Failed or cancelled mutations can
partly apply; inspect the surviving session before retrying.
See [macOS beta limitations](https://github.com/sbroenne/mcp-server-excel/blob/main/specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta)
and the [per-action inventory](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/MACOS-ACTION-INVENTORY.md).
Windows retains the complete backend.

[Documentation](https://excelmcpserver.dev/installation-mcp-server/) |
[Source](https://github.com/sbroenne/mcp-server-excel) |
[Issues](https://github.com/sbroenne/mcp-server-excel/issues)
