# Excel MCP Server - AI-Powered Excel Automation

[![GitHub](https://img.shields.io/badge/GitHub-sbroenne%2Fmcp--server--excel-blue)](https://github.com/sbroenne/mcp-server-excel)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)


**Control Microsoft Excel with AI through GitHub Copilot - just ask in natural language!**

**MCP Server for Excel** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands. Automate Power Query, DAX measures, VBA macros, PivotTables, Charts, formatting, and data transformations - no Excel programming knowledge required. 

**🛡️ 100% Safe - Uses Excel's Native COM API** - Zero risk of file corruption. Unlike third-party libraries that manipulate `.xlsx` files directly, this project uses Excel's official API ensuring complete safety and compatibility.

**💡 Interactive Development** - See results instantly in Excel. Create a query, run it, inspect the output, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing.

**🧪 LLM-Tested Quality** - Tool behavior validated with real AI agents using [agent-benchmark](https://github.com/mykhaliev/agent-benchmark). We test that LLMs correctly understand and use our tools.

## Features

The Excel MCP Server provides **22 specialized tools with 206 operations** for comprehensive Excel automation:

- 🔄 **Power Query** (1 tool, 10 ops) - Atomic workflows, M code management, load destinations
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
- �️ **Slicers** (1 tool, 8 ops) - Interactive filtering for PivotTables and Tables
- �🎨 **Conditional Formatting** (1 tool, 2 ops) - Rules and clearing

📚 **[Complete Feature Reference →](https://github.com/sbroenne/mcp-server-excel/blob/main/FEATURES.md)**

### Agent Skills (Bundled)

This extension includes **Agent Skills** following the [agentskills.io](https://agentskills.io) specification - providing domain-specific guidance for AI assistants. The skills enable GitHub Copilot to effectively understand Excel MCP Server capabilities, workflows, and best practices.

📚 **[View Agent Skills →](https://github.com/sbroenne/mcp-server-excel/blob/main/skills/excel-mcp/SKILL.md)**

**VS Code setup:** Enable the preview setting `chat.useAgentSkills` to allow Copilot to load skills. This extension installs the skill to `~/.copilot/skills/excel-mcp` for discovery.


## 💬 Example Prompts

**Create & Populate Data:**
- *"Create a new Excel file called SalesTracker.xlsx with a table for Date, Product, Quantity, Unit Price, and Total"*
- *"Put this data in A1:C4 - Name, Age, City / Alice, 30, Seattle / Bob, 25, Portland"*
- *"Add sample data and a formula column for Quantity times Unit Price"*

**Analysis & Visualization:**
- *"Create a PivotTable from this data showing total sales by Product, then add a bar chart"*
- *"Import products.csv with Power Query, load to Data Model, create a measure for Total Revenue"*
- *"Create a slicer for the Region field so I can filter the PivotTable interactively"*

**Formatting & Automation:**
- *"Format the Price column as currency and highlight values over $500 in green"*
- *"Export all Power Query M code to files for version control"*
- *"Show me Excel while you work"* - watch changes in real-time


## Quick Start

1. **Install this extension** (you just did!)
2. **Ask Copilot** in the chat panel:
   - "List all Power Query queries in workbook.xlsx"
   - "Create a DAX measure for year-over-year revenue growth"
   - "Export all Powere Queires and VBA modules to .vba files for version control"

**That's it!** The extension automatically installs .NET 10 runtime and includes a bundled MCP server.

➡️ **[Learn more and see examples](https://sbroenne.github.io/mcp-server-excel/)**

## Requirements

- **Windows OS** - Excel COM automation requires Windows
- **Microsoft Excel 2016+** - Must be installed on your system
- **.NET 10 Runtime** - **Automatically installed** by the extension

## Potential Issues

**"Excel is not installed" error:**
- Ensure Microsoft Excel 2016+ is installed on your Windows machine
- Try opening Excel manually to verify it works

**"VBA access denied" error:**
- VBA operations require one-time manual setup in Excel
- Go to: File → Options → Trust Center → Trust Center Settings → Macro Settings
- Check "Trust access to the VBA project object model"

**Copilot doesn't see Excel tools:**
- Restart VS Code after installing the extension
- ### Troubleshooting

- Check Output panel → "Excel MCP Server" for connection status

## Documentation & Support

- **[Complete Documentation](https://github.com/sbroenne/mcp-server-excel)** - Full guides and examples
- **[Report Issues](https://github.com/sbroenne/mcp-server-excel/issues)** - Bug reports and feature requests

## License & Privacy

MIT License - see [LICENSE](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)

Privacy Policy - see [PRIVACY.md](https://github.com/sbroenne/mcp-server-excel/blob/main/PRIVACY.md)

---

**Built with GitHub Copilot** | **Powered by Model Context Protocol**
