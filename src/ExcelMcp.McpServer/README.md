# ExcelMcp - Model Context Protocol Server for Excel

<!-- mcp-name: io.github.sbroenne/mcp-server-excel -->
mcp-name: io.github.sbroenne/mcp-server-excel

[![GitHub Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![GitHub Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=Downloads)](https://github.com/sbroenne/mcp-server-excel/releases)
[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue.svg)](https://github.com/sbroenne/mcp-server-excel)

**Control Excel with Natural Language** through AI assistants like GitHub Copilot, Claude, and ChatGPT. This MCP server enables AI-powered Excel automation for Power Query, DAX measures, VBA macros, PivotTables, Charts, and more.

➡️ **[Learn more and see examples](https://sbroenne.github.io/mcp-server-excel/)** 

**🛡️ 100% Safe - Uses Excel's Native COM API**

Unlike third-party libraries that manipulate `.xlsx` files (risking corruption), ExcelMcp uses **Excel's official COM automation API**. This guarantees zero risk of file corruption while you work interactively with live Excel files - see your changes happen in real-time.

**🔗 Unified Service Architecture** - The MCP Server forwards all requests to the shared ExcelMCP Service, enabling CLI and MCP to share sessions transparently.

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

📖 **Detailed setup instructions:** [Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md)

🎯 **Quick config examples:** [examples/mcp-configs/](https://github.com/sbroenne/mcp-server-excel/tree/main/examples/mcp-configs)

## 🛠️ What You Can Do

**25 specialized tools with 230 operations:**

- 🔄 **Power Query** (1 tool, 12 ops) - Atomic workflows, M code management, load destinations
- 📊 **Data Model/DAX** (2 tools, 19 ops) - Measures, relationships, model structure
- 🎨 **Excel Tables** (2 tools, 27 ops) - Lifecycle, filtering, sorting, structured references
- 📈 **PivotTables** (3 tools, 30 ops) - Creation, fields, aggregations, calculated members/fields
- 📉 **Charts** (2 tools, 29 ops) - Create, configure, series, formatting, data labels, trendlines
- 📝 **VBA** (1 tool, 6 ops) - Modules, execution, version control
- 📋 **Ranges** (4 tools, 46 ops) - Values, formulas, formatting, validation, protection
- 📄 **Worksheets** (2 tools, 16 ops) - Lifecycle, colors, visibility, cross-workbook moves
- 🔌 **Connections** (1 tool, 9 ops) - OLEDB/ODBC management and refresh
- 🏷️ **Named Ranges** (1 tool, 6 ops) - Parameters and configuration
- 📁 **Files** (1 tool, 6 ops) - Session management, workbook creation, IRM/AIP-protected file support
- 🧮 **Calculation Mode** (1 tool, 3 ops) - Get/set calculation mode and trigger recalculation
- 🎚️ **Slicers** (1 tool, 8 ops) - Interactive filtering for PivotTables and Tables
- 🎨 **Conditional Formatting** (1 tool, 2 ops) - Rules and clearing
- 📸 **Screenshot** (1 tool, 2 ops) - Capture ranges/sheets as PNG for visual verification
- 🪧 **Window Management** (1 tool, 9 ops) - Show/hide Excel, arrange, position, status bar feedback

📚 **[Complete Feature Reference →](../../FEATURES.md)** - Detailed documentation of all 230 operations

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

---

## 📋 Additional Resources

- **[GitHub Repository](https://github.com/sbroenne/mcp-server-excel)** - Source code, issues, discussions
- **[Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md)** - Detailed setup for all platforms
- **[VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)** - One-click installation
- **[CLI Documentation](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.CLI/README.md)** - Comprehensive commands for RPA and CI/CD automation

**License:** MIT  
**Privacy:** [PRIVACY.md](https://github.com/sbroenne/mcp-server-excel/blob/main/PRIVACY.md)  
**Platform:** Windows only (requires Excel 2016+)  
**Support:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
