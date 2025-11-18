# ExcelMcp - Model Context Protocol Server for Excel

<!-- mcp-name: io.github.sbroenne/mcp-server-excel -->
mcp-name: io.github.sbroenne/mcp-server-excel

[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue.svg)](https://github.com/sbroenne/mcp-server-excel)

**Control Excel with Natural Language** through AI assistants like GitHub Copilot, Claude, and ChatGPT. This MCP server enables AI-powered Excel automation for Power Query, DAX measures, VBA macros, PivotTables, and more.

➡️ **[Learn more and see examples](https://sbroenne.github.io/mcp-server-excel/)** 

Also includes a powerful CLI for RPA (Robotic Process Automation) and scripting workflows.

**🛡️ 100% Safe - Uses Excel's Native COM API**

Unlike third-party libraries that manipulate `.xlsx` files (risking corruption), ExcelMcp uses **Excel's official COM automation API**. This guarantees zero risk of file corruption while you work interactively with live Excel files - see your changes happen in real-time. Currently supports **163 operations across 12 specialized tools** with active development expanding capabilities.

**Requirements:** Windows OS + Excel 2016+

## 🚀 Installation

**Quick Setup Options:**

1. **VS Code Extension** - [One-click install](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) for GitHub Copilot
2. **Manual Install** - Works with Claude Desktop, Cursor, Cline, Windsurf, and other MCP clients
3. **MCP Registry** - Find us at [registry.modelcontextprotocol.io](https://registry.modelcontextprotocol.io/servers/io.github.sbroenne/mcp-server-excel)

**Manual Installation (All MCP Clients):**

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

**12 specialized MCP tools** for comprehensive Excel automation:

1. **Power Query & M Code** (11 actions) - Create, edit, optimize Power Query transformations with AI assistance
2. **Power Pivot / Data Model** (14 actions) - Build DAX measures, manage relationships, discover model structure, refresh data model
3. **Excel Tables** (23 actions) - Automate table creation, filtering, sorting, column management, structured references, number formatting
4. **PivotTables** (18 actions) - Create and configure PivotTables for interactive data analysis
5. **Ranges & Data** (43 actions) - Get/set values/formulas, number formatting, visual formatting (font, fill, border, alignment), data validation, find/replace, sort, insert/delete, hyperlinks, merge, conditional formatting, cell protection
6. **VBA Macros** (6 actions) - Import, update, run VBA code with version control integration
7. **Data Connections** (9 actions) - Manage OLEDB, ODBC, Text, Web connections and properties
8. **Worksheets** (16 actions) - Lifecycle management, tab colors, visibility controls
9. **Named Ranges** (7 actions) - Manage parameters and configuration through named ranges
10. **File Operations** (6 actions) - Create Excel workbooks (.xlsx/.xlsm), open/close workbook, save, test
11. **Conditional Formatting** (2 actions) - Add and clear conditional formatting rules

**Total: 11 tools with 155 operations**

**AI-Powered Workflows:**
- 💬 Natural language Excel commands through GitHub Copilot or Claude
- 🔄 Optimize Power Query M code for performance and readability  
- 📊 Build complex DAX measures with AI guidance
- 📋 Automate repetitive data transformations and formatting
- 🤖 **RPA:** Comprehensive CLI for robotic process automation, CI/CD, and batch processing


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
**Platform:** Windows only (requires Excel 2016+)  
**Support:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
