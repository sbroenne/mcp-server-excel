## ExcelMcp {{VERSION}}

### What's New
{{CHANGELOG}}

### Installation Options

**VS Code Extension** (Recommended)
- Search "ExcelMcp" in VS Code Marketplace and click Install
- Or download `excelmcp-{{VERSION}}.vsix` below
- Self-contained: no .NET runtime or SDK required
- Includes both MCP Server and CLI (`excelcli`)
- Agent skills (excel-mcp + excel-cli) registered automatically via `chatSkills`

**Claude Desktop (MCPB)**
- Download `excel-mcp-{{VERSION}}.mcpb` and double-click to install

**Standalone Executables** (Primary — no .NET runtime required)
- MCP Server: Download `ExcelMcp-MCP-Server-{{VERSION}}-windows.zip`, extract `mcp-excel.exe`
- CLI: Download `ExcelMcp-CLI-{{VERSION}}-windows.zip`, extract `excelcli.exe`
- Add the exe(s) to your PATH, then configure your MCP client with command `mcp-excel`

**NuGet (.NET Tool)** (Secondary — requires .NET 10 runtime)
```powershell
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

**Agent Skills** (for AI coding assistants)
- VS Code Extension includes both skills automatically (excel-mcp + excel-cli)
- Install via Skills CLI: `npx skills add sbroenne/mcp-server-excel --skill excel-cli` or `--skill excel-mcp`
- Or download `excel-skills-v{{VERSION}}.zip`

### Requirements
- Windows OS
- Microsoft Excel 2016+
- No .NET runtime required for VS Code Extension, MCPB, or standalone executables
- .NET 10 Runtime required for NuGet (.NET tool) installation only

### Documentation
- [Website](https://excelmcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- [Changelog](https://github.com/sbroenne/mcp-server-excel/blob/main/CHANGELOG.md)
