# ExcelMcp MCP Server

Run the self-contained ExcelMcp server through npm:

```powershell
npx -y @sbroenne/mcp-server-excel
```

The package is Windows-only and requires Microsoft Excel 2016 or later. It does
not require the .NET SDK or a separately installed .NET runtime.

The Node.js entry point only launches the packaged .NET server. MCP tools and
Excel automation continue to run in the existing ExcelMcp implementation.

[Documentation](https://excelmcpserver.dev/installation-mcp-server/) |
[Source](https://github.com/sbroenne/mcp-server-excel) |
[Issues](https://github.com/sbroenne/mcp-server-excel/issues)
