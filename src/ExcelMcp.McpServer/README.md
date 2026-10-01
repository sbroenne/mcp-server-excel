# ExcelMcp - Model Context Protocol Server for Excel

<!-- mcp-name: io.github.sbroenne/mcp-server-excel -->
mcp-name: io.github.sbroenne/mcp-server-excel

[![GitHub Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![GitHub Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=Downloads)](https://github.com/sbroenne/mcp-server-excel/releases)
[![npm](https://img.shields.io/npm/v/%40sbroenne%2Fmcp-server-excel)](https://www.npmjs.com/package/@sbroenne/mcp-server-excel)
[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue.svg)](https://github.com/sbroenne/mcp-server-excel)

**Control Excel with Natural Language** through AI assistants like GitHub Copilot, Claude, and ChatGPT. This MCP server enables AI-powered Excel automation for Power Query, DAX measures, VBA macros, PivotTables, Charts, and more.

➡️ **[Learn more and see examples](https://excelmcpserver.dev/)** 

**⚡ Powered by the Real Excel Engine**

Unlike file-parser libraries, ExcelMcp drives the **actual Excel application**. Windows uses the complete COM backend. macOS uses a capability-gated Apple Events backend for the documented initial operation set; unsupported operations fail explicitly.

**🔗 In-Process Service Architecture** - The MCP Server hosts the ExcelMcp Service in-process and calls it directly (no pipe), for low-latency Excel automation. The CLI is an equal entry point that runs the same service as a background daemon.

**CLI also available:** `mcp-excel` (MCP Server) and `excelcli` (CLI) are distributed as standalone self-contained executables — no .NET runtime required.

**Requirements:** Windows 10+ with Excel 2016+, or Apple Silicon macOS with Excel for Mac 16.112+

> **macOS support is experimental beta.** Power Query, VBA, Data Model/DAX/OLAP,
> Tables, PivotTables, charts, slicers, connections, QueryTables, XML Maps,
> screenshots, advanced visual formatting, and Python result reads are not
> supported. Windows retains the complete backend. See
> [macOS beta limitations](../../specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta);
> test on copies of important workbooks.

## 🚀 Installation

**Quick Setup Options:**

1. **VS Code Extension** - [One-click install](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) for GitHub Copilot
2. **npm** - Run through `npx` with no .NET installation or manual executable setup
3. **Standalone exe** - Works with Claude Desktop, Cursor, Cline, Windsurf, and other MCP clients
4. **MCP Registry** - Find us at [registry.modelcontextprotocol.io](https://registry.modelcontextprotocol.io/v0/servers?search=io.github.sbroenne/mcp-server-excel) as `io.github.sbroenne/mcp-server-excel`

**Manual Installation (All MCP Clients):**

**Primary — npm (no .NET runtime required):**

```powershell
npx -y @sbroenne/mcp-server-excel@latest
```

Configure MCP clients with `command: "npx"` and
`args: ["-y", "@sbroenne/mcp-server-excel@latest"]`. Requires Node.js 18+;
npm resolves `@latest` at launch using its normal cache policy.

**Standalone executable:**

```powershell
# Download from GitHub Releases:
# Windows: ExcelMcp-MCP-Server-{version}-windows.zip → extract mcp-excel.exe
# macOS ARM64: ExcelMcp-MCP-Server-{version}-macos-arm64.zip → extract mcp-excel
```

**Secondary — .NET Global Tool (requires .NET 10 runtime):**

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
```

**Supported AI Assistants:**
- ✅ GitHub Copilot (VS Code, Visual Studio)
- ✅ Claude Desktop
- ✅ Cursor
- ✅ Cline (VS Code Extension)
- ✅ Windsurf
- ✅ Any MCP-compatible client

📖 **Detailed setup instructions:** [MCP Server Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION-MCP-SERVER.md)

🎯 **Quick config examples:** [examples/mcp-configs/](https://github.com/sbroenne/mcp-server-excel/tree/main/examples/mcp-configs)

## 🛠️ What You Can Do

**31 specialized tools with 326 operations** are available through the complete
Windows backend. Apple Silicon macOS exposes only actions marked enabled in the
[generated capability inventory](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/MACOS-ACTION-INVENTORY.md);
the enabled set includes workbook lifecycle, worksheet management, values and
formulas, number formats, row and column sizing, merged cells, cell locking,
named ranges, Goal Seek, Data Tables, calculation, and Python formula writes.
Unavailable operations return `PlatformNotSupported`.

📚 **[Complete Feature Reference →](https://github.com/sbroenne/mcp-server-excel/blob/main/FEATURES.md)** - Detailed documentation of all 326 operations, grouped by category

**AI-Powered Workflows:**
The Power Query, DAX, and Show Excel/Agent Mode examples require Windows.
- 💬 Natural language Excel commands through GitHub Copilot, Claude, or ChatGPT
- 🔄 Optimize Power Query M code for performance and readability  
- 📊 Build complex DAX measures with AI guidance
- 📋 Automate repetitive data transformations and formatting
- 👀 **Show Excel Mode** - Say "Show me Excel while you work" to watch changes live


---

## Session workflow

List and match the intended workbook, reuse its session or open/create, perform
the work, then list and check its `canClose`. Close only when authorized and
choose `save: true` or `save: false` explicitly. Closing without saving discards
all unsaved edits and has no tool-level undo.

MCP inputs, open/create results, list entries, and session error context use
`session_id`. The legacy `sessionId` input is rejected. CLI JSON keeps its
`sessionId` convention; its sessions are separate.

Calls within a session execute one at a time, but concurrent requests and
responses have no guaranteed order. Wait for dependent calls. Writes attempt to
restore the calculation mode; restoration can fail without failing the write.
Use `get-mode` when subsequent work depends on the mode. Manual mode requires
explicit calculation before relying on dependent values.

## 💡 Example Use Cases

Plain worksheet values/formulas work on both platforms. PivotTables, charts,
Power Query/Data Model, and slicers below require Windows in the Mac beta.

**"Create a sales tracker with Date, Product, Quantity, Unit Price, and Total columns"**  
→ AI creates the workbook, adds headers, enters sample data, and builds formulas

**"Create a PivotTable from this data showing total sales by Product, then add a chart"**  
→ AI creates PivotTable, configures fields, and adds a linked visualization

**"Import products.csv with Power Query, load to Data Model, create a Total Revenue measure"**  
→ AI imports data, adds to Power Pivot, and creates DAX measures for analysis

**"Create a slicer for the Region field so I can filter interactively"**  
→ AI adds slicers connected to PivotTables or Tables for point-and-click filtering

**"Put this data in A1: Name, Age / Alice, 30 / Bob, 25"**  
→ AI writes data directly to cells using natural delimiters you provide

---

## 📋 Additional Resources

- **[GitHub Repository](https://github.com/sbroenne/mcp-server-excel)** - Source code, issues, discussions
- **[MCP Server Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION-MCP-SERVER.md)** - Detailed setup for all platforms
- **[VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)** - One-click installation
- **[CLI Documentation](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.CLI/README.md)** - Comprehensive commands for RPA and CI/CD automation

**License:** MIT  
**Privacy:** [PRIVACY.md](https://github.com/sbroenne/mcp-server-excel/blob/main/PRIVACY.md)
**Platform:** Windows x64 (complete backend) and Apple Silicon macOS
(experimental beta, capability-gated backend).
**Support:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
