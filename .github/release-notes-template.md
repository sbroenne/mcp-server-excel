## ExcelMcp {{VERSION}}

### What's New
{{CHANGELOG}}

### Installation Options

**VS Code Extension** (Recommended)
- Search "ExcelMcp" in VS Code Marketplace and click Install
- Or download the `win32-x64` or `darwin-arm64` VSIX below
- Self-contained: no .NET runtime or SDK required
- Includes the MCP Server and `excel-mcp` skill; install the CLI separately

**Claude Desktop (MCPB)**
- Download the matching Windows or macOS ARM64 MCPB and double-click to install

**npm MCP Server** (Primary — no .NET runtime required)
```powershell
npx -y @sbroenne/mcp-server-excel
```

**npm CLI** (Primary — no .NET runtime required)
```powershell
npx -y @sbroenne/excelcli --help
# Or install the command on your PATH:
npm install --global @sbroenne/excelcli
```

**Standalone Executables** (no .NET runtime required)
- MCP Server: Download `ExcelMcp-MCP-Server-{{VERSION}}-windows.zip`, extract `mcp-excel.exe`
- CLI: Download `ExcelMcp-CLI-{{VERSION}}-windows.zip`, extract `excelcli.exe`
- macOS uses the `*-macos-arm64.zip` archive
- Add the executable(s) to your PATH, then configure your MCP client with command `mcp-excel`

**NuGet (.NET Tool)** (Secondary — requires .NET 10 runtime)
```powershell
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

**Agent Skills** (for AI coding assistants)
- VS Code Extension includes the `excel-mcp` skill; install `excel-cli` separately
- Install via Skills CLI: `npx skills add sbroenne/mcp-server-excel --skill excel-cli` or `--skill excel-mcp`
- Or download `excel-skills-v{{VERSION}}.zip`

### Requirements
- Windows x64 with Microsoft Excel 2016+, or Apple Silicon macOS with Excel for Mac 16.112+
- Node.js 18+ for npm or Skills CLI installation
- No .NET runtime required for npm, VS Code Extension, MCPB, or standalone executables
- .NET 10 Runtime required for NuGet (.NET tool) installation only

### Documentation
- [Website](https://excelmcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- [Changelog](https://github.com/sbroenne/mcp-server-excel/blob/main/CHANGELOG.md)
