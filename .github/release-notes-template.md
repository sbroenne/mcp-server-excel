## ExcelMcp {{VERSION}}

### What's New
{{CHANGELOG}}

### Installation Options

**Apple Silicon macOS support is experimental beta**, not full Windows parity.
Power Query, VBA, Data Model/DAX/OLAP, Tables, PivotTables, charts, slicers,
connections, QueryTables, XML Maps, screenshots, advanced visual formatting,
and Python result reads are unsupported on Mac. See
[the beta support reference](https://github.com/sbroenne/mcp-server-excel/blob/main/specs/MACOS-SUPPORT.md);
use copies of important workbooks.

**VS Code Extension** (Recommended)
- Search "ExcelMcp" in VS Code Marketplace and click Install
- Or download `excel-mcp-{{VERSION}}.vsix`,
  `excel-mcp-{{VERSION}}-win32-arm64.vsix`, or
  `excelmcp-{{VERSION}}-darwin-arm64.vsix` below
- Self-contained: no .NET runtime or SDK required
- Includes the MCP Server and its `excel-mcp` skill; install the CLI separately
- The MCP skill is registered automatically via `chatSkills`

**Claude Desktop (MCPB)**
- Download `excel-mcp-{{VERSION}}.mcpb` and double-click to install
- Requires Node.js/npm with npx on PATH; launches the npm server with `@latest`
- Older binary MCPB installations need a one-time replacement with this bundle
- On Apple Silicon macOS, download
  `excel-mcp-{{VERSION}}-macos-arm64.mcpb`; it contains the native beta runtime

**npm MCP Server** (Primary — no .NET runtime required)
```powershell
npx -y @sbroenne/mcp-server-excel@latest
```

**npm CLI** (Primary — no .NET runtime required)
```powershell
npx -y @sbroenne/excelcli@latest --help
# Or install the command on your PATH:
npm install --global @sbroenne/excelcli@latest
```

**Standalone Executables** (no .NET runtime required)
- MCP Server: Download `ExcelMcp-MCP-Server-{{VERSION}}-windows.zip`, extract `mcp-excel.exe`
- CLI: Download `ExcelMcp-CLI-{{VERSION}}-windows.zip`, extract `excelcli.exe`
- Mac: use `ExcelMcp-MCP-Server-{{VERSION}}-macos-arm64.zip` (`mcp-excel`) or
  `ExcelMcp-CLI-{{VERSION}}-macos-arm64.zip` (`excelcli`)
- Add the executable(s) to PATH, then configure your MCP client with command `mcp-excel`

**NuGet (.NET Tool)** (Secondary — requires .NET 10 runtime)
```powershell
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

**Agent Skills** (optional report formatting)
- VS Code Extension registers `excel-mcp-report-formatting`
- Install via Skills CLI: `npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli-report-formatting` or `--skill excel-mcp-report-formatting`
- Or download `excel-skills-v{{VERSION}}.zip`

### Requirements
- Windows x64 with desktop Excel 2016+, or Apple Silicon macOS with Excel for
  Mac 16.112+ (experimental beta)
- An interactive desktop session; Intel Macs, Linux, and headless hosts are unsupported
- Node.js 18+ with npm/npx for npm and the Windows MCPB installation
- No .NET runtime required for npm, VS Code Extension, MCPB, or standalone executables
- .NET 10 Runtime required for NuGet (.NET tool) installation only

`@latest` is resolved when launching, subject to normal npm caching. It does not
replace a running server or CLI service. Finish and explicitly save/close
workbook sessions before restarting to use an updated runtime.

### Documentation
- [Website](https://excelmcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- [Changelog](https://github.com/sbroenne/mcp-server-excel/blob/main/CHANGELOG.md)
