# ExcelMcp - Model Context Protocol Server for Excel

<!-- mcp-name: io.github.sbroenne/mcp-server-excel -->
mcp-name: io.github.sbroenne/mcp-server-excel

[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue.svg)](https://github.com/sbroenne/mcp-server-excel)

**Control Excel with Natural Language** through AI assistants like GitHub Copilot, Claude, and ChatGPT. This MCP server enables AI-powered Excel automation for Power Query, DAX measures, VBA macros, PivotTables, Charts, and more.

➡️ **[Learn more and see examples](https://sbroenne.github.io/mcp-server-excel/)** 

**🛡️ 100% Safe - Uses Excel's Native COM API**

Unlike third-party libraries that manipulate `.xlsx` files (risking corruption), ExcelMcp uses **Excel's official COM automation API**. This guarantees zero risk of file corruption while you work interactively with live Excel files - see your changes happen in real-time.

**🔗 Unified Service Architecture** - The MCP Server forwards all requests to the shared ExcelMCP Service, enabling CLI and MCP to share sessions transparently.

**📦 Unified Package - Includes CLI:** This package includes both the MCP Server (`mcp-excel`) and CLI (`excelcli`). Install once, get both tools! For advanced users who prefer command-line scripting, the CLI interface supports RPA workflows, CI/CD pipelines, and batch automation. CLI has 16 command categories with 216 operations matching the MCP Server (24 tools with 216 operations).

**Requirements:** Windows OS + Excel 2016+

## 🚀 Installation

**Quick Setup Options:**

1. **VS Code Extension** - [One-click install](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) for GitHub Copilot
2. **Manual Install** - Works with Claude Desktop, Cursor, Cline, Windsurf, and other MCP clients
3. **MCP Registry** - Find us at [registry.modelcontextprotocol.io](https://registry.modelcontextprotocol.io/servers/io.github.sbroenne/mcp-server-excel)

**Manual Installation (All MCP Clients):**

Requires .NET 10 Runtime or SDK

```powershell
# Install unified package (includes MCP Server + CLI)
dotnet tool install --global Sbroenne.ExcelMcp.McpServer

# Both tools are now available:
# - mcp-excel (MCP Server for AI assistants)
# - excelcli (CLI for coding agents and scripting)
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

**24 specialized tools with 216 operations:**

- 🔄 **Power Query** (1 tool, 11 ops) - Atomic workflows, M code management, load destinations
- 📊 **Data Model/DAX** (2 tools, 18 ops) - Measures, relationships, model structure
- 🎨 **Excel Tables** (2 tools, 27 ops) - Lifecycle, filtering, sorting, structured references
- 📈 **PivotTables** (3 tools, 30 ops) - Creation, fields, aggregations, calculated members/fields
- 📉 **Charts** (2 tools, 26 ops) - Create, configure, series, formatting, data labels, trendlines
- 📝 **VBA** (1 tool, 6 ops) - Modules, execution, version control
- 📋 **Ranges** (4 tools, 42 ops) - Values, formulas, formatting, validation, protection
- 📄 **Worksheets** (2 tools, 16 ops) - Lifecycle, colors, visibility, cross-workbook moves
- 🔌 **Connections** (1 tool, 9 ops) - OLEDB/ODBC management and refresh
- 🏷️ **Named Ranges** (1 tool, 6 ops) - Parameters and configuration
- 📁 **Files** (1 tool, 6 ops) - Session management and workbook creation
- 🧮 **Calculation Mode** (1 tool, 3 ops) - Get/set calculation mode and trigger recalculation
- 🎚️ **Slicers** (1 tool, 8 ops) - Interactive filtering for PivotTables and Tables
- 🎨 **Conditional Formatting** (1 tool, 2 ops) - Rules and clearing
- 📸 **Screenshot** (1 tool, 2 ops) - Capture ranges/sheets as PNG for visual verification

📚 **[Complete Feature Reference →](../../FEATURES.md)** - Detailed documentation of all 216 operations

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
