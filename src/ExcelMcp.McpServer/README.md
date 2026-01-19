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

**Optional CLI Tool:** For advanced users who prefer command-line scripting, ExcelMcp includes a CLI interface for RPA workflows, CI/CD pipelines, and batch automation. CLI has 13 command categories with 187 operations matching the MCP Server (21 tools with 187 operations).

**Requirements:** Windows OS + Excel 2016+

## 🚀 Installation

**Quick Setup Options:**

1. **VS Code Extension** - [One-click install](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) for GitHub Copilot
2. **Manual Install** - Works with Claude Desktop, Cursor, Cline, Windsurf, and other MCP clients
3. **MCP Registry** - Find us at [registry.modelcontextprotocol.io](https://registry.modelcontextprotocol.io/servers/io.github.sbroenne/mcp-server-excel)

**Manual Installation (All MCP Clients):**

Requires .NET 10 Runtime or SDK

```powershell
# Install MCP Server
dotnet tool install --global Sbroenne.ExcelMcp.McpServer

# Configure your AI assistant
# See examples/mcp-configs/ for ready-to-use configs
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

**21 specialized tools with 187 operations:**

- 🔄 **Power Query** (1 tool, 10 ops) - Atomic workflows, M code management, load destinations
- 📊 **Data Model/DAX** (2 tools, 17 ops) - Measures, relationships, model structure
- 🎨 **Excel Tables** (2 tools, 24 ops) - Lifecycle, filtering, sorting, structured references
- 📈 **PivotTables** (3 tools, 30 ops) - Creation, fields, aggregations, calculated members/fields
- 📉 **Charts** (2 tools, 14 ops) - Create, configure, manage series and formatting
- 📝 **VBA** (1 tool, 6 ops) - Modules, execution, version control
- 📋 **Ranges** (4 tools, 42 ops) - Values, formulas, formatting, validation, protection
- 📄 **Worksheets** (2 tools, 16 ops) - Lifecycle, colors, visibility, cross-workbook moves
- 🔌 **Connections** (1 tool, 9 ops) - OLEDB/ODBC management and refresh
- 🏷️ **Named Ranges** (1 tool, 6 ops) - Parameters and configuration
- 📁 **Files** (1 tool, 6 ops) - Session management and workbook creation
- 🎨 **Conditional Formatting** (1 tool, 2 ops) - Rules and clearing

📚 **[Complete Feature Reference →](../../FEATURES.md)** - Detailed documentation of all 187 operations

**AI-Powered Workflows:**
- 💬 Natural language Excel commands through GitHub Copilot, Claude, or ChatGPT
- 🔄 Optimize Power Query M code for performance and readability  
- 📊 Build complex DAX measures with AI guidance
- 📋 Automate repetitive data transformations and formatting
- 👀 **Show Excel Mode** - Say "Show me Excel while you work" to watch changes live


---

## 💡 Example Use Cases

**"Optimize all my Power Queries in this workbook"**  
→ AI analyzes M code, optimizes query folding, improves step organization

**"Create a PivotTable showing sales by region and product"**  
→ AI creates PivotTable, adds fields, sets aggregations, applies formatting

**"Format revenue columns as currency with bold headers and blue background"**  
→ AI applies number formatting, font styles, and cell colors

**"Build a DAX measure for year-over-year growth"**  
→ AI writes DAX formula, sets currency format, adds to Data Model

**"Export all Power Query M code to Git for version control"**  
→ CLI batch exports all queries to .pq files for source control workflows

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
