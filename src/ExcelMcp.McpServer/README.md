# ExcelMcp - Model Context Protocol Server for Excel

<!-- mcp-name: io.github.sbroenne/mcp-server-excel -->
mcp-name: io.github.sbroenne/mcp-server-excel

[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue.svg)](https://github.com/sbroenne/mcp-server-excel)

**Control Excel with Natural Language** through AI assistants like GitHub Copilot, Claude, and ChatGPT. This MCP server enables AI-powered Excel automation for Power Query, DAX measures, VBA macros, PivotTables, and more.

**🛡️ 100% Safe - Uses Excel's Native COM API**

Unlike third-party libraries that manipulate `.xlsx` files (risking corruption), ExcelMcp uses **Excel's official COM automation API**. This guarantees zero risk of file corruption while you work interactively with live Excel files - see your changes happen in real-time. Currently supports 80+ operations with active development expanding capabilities.

**Requirements:** Windows OS + Excel 2016+

**Installation: Global .NET Tool**

```powershell
# Install .NET 8 SDK
winget install Microsoft.DotNet.SDK.8

# Install ExcelMcp MCP server
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
```

**Configure AI Assistant** - Add to your MCP configuration:

```json
{
  "servers": {
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "mcp-excel"]
    }
  }
}
```

## 🛠️ What You Can Do

**11 specialized MCP tools** for comprehensive Excel automation:

1. **Power Query & M Code** (11 actions) - Create, edit, refactor Power Query transformations with AI assistance
2. **Power Pivot / Data Model** (14 actions) - Build DAX measures, manage relationships, export to .dax files for Git workflows
3. **Excel Tables** (22 actions) - Automate table creation, filtering, sorting, column management, structured references
4. **PivotTables** (20 actions) - Create and configure PivotTables for interactive data analysis
5. **Ranges & Data** (38+ actions) - Get/set values/formulas, number formatting, visual formatting (font, fill, border, alignment), data validation, find/replace, sort, insert/delete, hyperlinks
6. **VBA Macros** (7 actions) - Export, import, run VBA code with version control integration
7. **Data Connections** (11 actions) - Manage OLEDB, ODBC, Text, Web connections and properties
8. **Worksheets** (5 actions) - Create, rename, copy, delete worksheets
9. **Named Ranges** (6 actions) - Manage parameters and configuration through named ranges
10. **File Operations** (1 action) - Create Excel workbooks (.xlsx/.xlsm)
11. **Batch Sessions** (3 actions) - Group multiple operations for better performance

**AI-Powered Workflows:**
- 💬 Natural language Excel commands through GitHub Copilot or Claude
- 🔄 Refactor Power Query M code for performance and readability  
- 📊 Build complex DAX measures with AI guidance
- 🧪 Enhance VBA macros with error handling and best practices
- 📋 Automate repetitive data transformations and analysis


---

## 💡 Example Use Cases

**"Refactor this slow Power Query"**  
→ AI analyzes M code, optimizes query folding, improves step organization

**"Add error handling to my VBA macro"**  
→ AI exports code, adds Try-Catch patterns, implements logging, updates module

**"Create a PivotTable showing sales by region and product"**  
→ AI creates PivotTable, adds fields, sets aggregations, applies formatting

**"Build a DAX measure for year-over-year growth"**  
→ AI writes DAX formula, sets currency format, adds to Data Model

---

## 📋 Additional Resources

- **[GitHub Repository](https://github.com/sbroenne/mcp-server-excel)** - Source code, issues, discussions
- **[Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md)** - Detailed setup for all platforms
- **[VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excelmcp)** - One-click installation
- **[CLI Documentation](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/CLI.md)** - 50+ commands for scripting

**License:** MIT  
**Platform:** Windows only (requires Excel 2016+)  
**Support:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
