# ExcelMcp - Model Context Protocol Server for Excel

<!-- mcp-name: io.github.sbroenne/mcp-server-excel -->
mcp-name: io.github.sbroenne/mcp-server-excel

[![GitHub Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![GitHub Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=Downloads)](https://github.com/sbroenne/mcp-server-excel/releases)
[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue.svg)](https://github.com/sbroenne/mcp-server-excel)

<!--start-->
**Control Excel with Natural Language** through AI assistants like GitHub Copilot, Claude, and ChatGPT. This MCP server enables AI-powered Excel automation for Power Query, DAX measures, VBA macros, PivotTables, Charts, and more.

➡️ **[Learn more and see examples](https://excelmcpserver.dev/)** 

**⚡ Powered by the Real Excel Engine**

Unlike file-parser libraries that rewrite `.xlsx` files directly, ExcelMcp drives the **actual Excel application** through its official COM API. That means it can run live operations file-based tools can't — refresh Power Query, recalculate, refresh PivotTables and the Data Model, evaluate DAX, run VBA and Python `=PY()` — and edit your existing workbooks with formulas, PivotTables, charts, macros and formatting left intact. Watch it happen in real time.

**🔗 In-Process Service Architecture** - The MCP Server hosts the ExcelMCP Service in-process and calls it directly (no pipe), for low-latency Excel automation. The CLI is an equal entry point that runs the same service as a background daemon.

**CLI also available:** `mcp-excel.exe` (MCP Server) and `excelcli.exe` (CLI) are distributed as standalone self-contained executables — no .NET runtime required.

**Requirements:** Windows OS + Excel 2016+

## 🚀 Installation

**Quick Setup Options:**

1. **VS Code Extension** - [One-click install](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) for GitHub Copilot
2. **Standalone exe** - Works with Claude Desktop, Cursor, Cline, Windsurf, and other MCP clients
3. **MCP Registry** - Find us at [registry.modelcontextprotocol.io](https://registry.modelcontextprotocol.io/servers/io.github.sbroenne/mcp-server-excel)

**Manual Installation (All MCP Clients):**

**Primary — Standalone exe (no .NET runtime required):**

```powershell
# Download from latest release:
# https://github.com/sbroenne/mcp-server-excel/releases/latest
# ExcelMcp-MCP-Server-{version}-windows.zip → extract mcp-excel.exe

# Add to PATH, then configure your MCP client:
# { "command": "mcp-excel" }
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

📖 **Detailed setup instructions:** [MCP Server Installation Guide](https://excelmcpserver.dev/installation-mcp-server/)

🎯 **Quick config examples:** [examples/mcp-configs/](https://github.com/sbroenne/mcp-server-excel/tree/main/examples/mcp-configs)

## 🛠️ What You Can Do

**26 specialized tools with 232 operations** covering Power Query, Data Model/DAX, PivotTables, Excel Tables, Charts, VBA, Ranges, Worksheets, Connections, Named Ranges, File/Session management, Calculation Mode, Slicers, Conditional Formatting, Screenshots, and Window Management.

📚 **[Complete Feature Reference →](https://excelmcpserver.dev/features/)** - Detailed documentation of all 232 operations, grouped by category

**AI-Powered Workflows:**
- 💬 Natural language Excel commands through GitHub Copilot, Claude, or ChatGPT
- 🔄 Optimize Power Query M code for performance and readability  
- 📊 Build complex DAX measures with AI guidance
- 📋 Automate repetitive data transformations and formatting
- 👀 **Show Excel Mode** - Say "Show me Excel while you work" to watch changes live


---

## 💡 Example Use Cases

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

<!--end-->

---

## 📋 Additional Resources

- **[GitHub Repository](https://github.com/sbroenne/mcp-server-excel)** - Source code, issues, discussions
- **[MCP Server Installation Guide](https://excelmcpserver.dev/installation-mcp-server/)** - Detailed setup for all platforms
- **[VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)** - One-click installation
- **[CLI Documentation](https://excelmcpserver.dev/cli/)** - Comprehensive commands for RPA and CI/CD automation

**License:** MIT  
**Privacy:** [PRIVACY.md](https://github.com/sbroenne/mcp-server-excel/blob/main/PRIVACY.md)  
**Platform:** Windows only (requires Excel 2016+)  
**Support:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
