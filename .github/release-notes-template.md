## ExcelMcp {{VERSION}}

### What's New
{{CHANGELOG}}

### Installation Options

**VS Code Extension** (Recommended)
- Search "ExcelMcp" in VS Code Marketplace and click Install
- Or download `excel-mcp-{{VERSION}}.vsix` below
- Self-contained: no .NET runtime or SDK required
- Includes the MCP Server and `excel-mcp` skill; install the CLI separately

**Claude Desktop (MCPB)**
- Download `excel-mcp-{{VERSION}}.mcpb` and double-click to install
- Requires Node.js/npm with npx on PATH; launches the npm server with `@latest`
- Older binary MCPB installations need a one-time replacement with this bundle

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
- Add the exe(s) to your PATH, then configure your MCP client with command `mcp-excel`

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
- Windows OS
- Microsoft Excel 2016+
- Interactive Windows desktop with Excel available to the signed-in user
- Node.js 18+ with npm/npx for npm and MCPB installation
- No .NET runtime required for npm, VS Code Extension, MCPB, or standalone executables
- .NET 10 Runtime required for NuGet (.NET tool) installation only

`@latest` is resolved when launching, subject to normal npm caching. It does not
replace a running server or CLI service. Finish and explicitly save/close
workbook sessions before restarting to use an updated runtime.

### Documentation
- [Website](https://excelmcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- [Changelog](https://github.com/sbroenne/mcp-server-excel/blob/main/CHANGELOG.md)
