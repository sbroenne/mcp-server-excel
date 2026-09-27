# ExcelMcp MCP Server

Run the self-contained ExcelMcp server through npm:

```powershell
npx -y @sbroenne/mcp-server-excel
```

The package supports Windows x64/Arm64 (through the x64 runtime) with Microsoft
Excel 2016 or later and macOS x64/Arm64 with Excel for Mac 16.112 or later.
Intel macOS uses the separate `darwin-x64` runtime package; physical Intel Mac
Excel execution remains unverified. It does not require the .NET SDK or a
separately installed .NET runtime. Keep optional dependencies enabled so npm installs the
matching native runtime package.

The Node.js entry point only launches the packaged .NET server. MCP tools and
Excel automation continue to run in the existing ExcelMcp implementation.

[Documentation](https://excelmcpserver.dev/installation-mcp-server/) |
[Source](https://github.com/sbroenne/mcp-server-excel) |
[Issues](https://github.com/sbroenne/mcp-server-excel/issues)
