# ExcelMcp MCP Server

Run the self-contained ExcelMcp server through npm:

```powershell
npx -y @sbroenne/mcp-server-excel@latest
```

The package requires Node.js 18+, Windows, Microsoft Excel 2016 or later, and an
interactive desktop. It does
not require the .NET SDK or a separately installed .NET runtime.
ARM64 Node.js selects the native ARM64 runtime; x64 Node.js selects x64,
including on ARM64 Windows. Keep optional dependencies enabled. A missing
matching package is an error, not a fallback to another architecture.

The Node.js entry point only launches the packaged .NET server. MCP tools and
Excel automation continue to run in the existing ExcelMcp implementation.

`@latest` selects the current npm release at startup using normal npm caching.
Network access is needed for downloads and update checks. Restart your MCP
server after safely finishing workbook work to run an updated version.

[Documentation](https://excelmcpserver.dev/installation-mcp-server/) |
[Source](https://github.com/sbroenne/mcp-server-excel) |
[Issues](https://github.com/sbroenne/mcp-server-excel/issues)
