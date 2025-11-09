# Excel MCP Server - AI-Powered Excel Automation

**Control Microsoft Excel with Natural Language through AI assistants like GitHub Copilot and Claude.**

## What is Excel MCP Server?

Excel MCP Server is a Model Context Protocol (MCP) server that enables AI assistants to automate Microsoft Excel. Talk to Excel in plain English to automate Power Query, DAX measures, VBA macros, PivotTables, formatting, and data transformations - no programming required.

## Key Features

- **165 Operations** across 11 specialized tools
- **100% Safe** - Uses Excel's native COM API (zero corruption risk)
- **Real-time Interaction** - See changes happen live in Excel
- **Natural Language Control** - Describe what you want, AI does the rest
- **Comprehensive Automation**: Power Query, Power Pivot, VBA, Tables, ranges, formatting

## Use Cases

### Data Transformation & Analysis
- Optimize Power Query M code for performance
- Create PivotTables from data with natural language
- Build DAX measures with AI guidance
- Transform and clean data automatically

### Formatting & Styling (No Programming)
- Format columns as currency, dates, percentages
- Apply conditional formatting with color rules
- Create Excel Tables with auto-filters
- Add data validation dropdowns

### Workflow Automation
- Find and replace across ranges
- Sort and filter data programmatically
- Merge cells and auto-fit columns
- Manage VBA macros with version control

## Installation

### Quick Start (VS Code + GitHub Copilot)
Install the [Excel MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excelmcp) for one-click setup.

### Manual Installation (Any MCP Client)
```json
{
  "mcpServers": {
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "excel-mcp"]
    }
  }
}
```

## Requirements

- **Windows OS** (Excel COM API requires Windows)
- **Microsoft Excel 2016 or later**
- **.NET 8.0 Runtime** (automatically installed with NuGet tool)

## Documentation

- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- [Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md)
- [Contributing Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/CONTRIBUTING.md)
- [API Reference](https://github.com/sbroenne/mcp-server-excel/tree/main/docs)

## Examples

### Example 1: Optimize Power Query
```
You: "This Power Query is taking 5 minutes to refresh. Can you optimize it?"

AI analyzes your M code, identifies inefficiencies, and applies best practices automatically.
```

### Example 2: Create PivotTable
```
You: "Create a PivotTable from SalesData showing top 10 products by region"

AI creates the PivotTable with proper field configuration in seconds.
```

### Example 3: Format Data
```
You: "Format revenue as currency, make headers bold blue, and add borders"

AI applies all formatting directly to your Excel file.
```

## Why Excel MCP Server?

- **Safe**: Official COM API - zero corruption risk
- **Interactive**: See changes in real-time
- **Comprehensive**: 165 operations cover most Excel tasks
- **AI-Powered**: Natural language control
- **Developer-Friendly**: Built-in CLI for RPA and scripting

## Related Projects

- [Model Context Protocol](https://modelcontextprotocol.io/) - The protocol specification
- [MCP Servers](https://github.com/modelcontextprotocol/servers) - Official MCP server examples
- [Claude Desktop](https://claude.ai/download) - AI assistant with MCP support
- [GitHub Copilot](https://github.com/features/copilot) - AI coding assistant

## Keywords

Excel automation, MCP server, Model Context Protocol, GitHub Copilot, Claude AI, Power Query, M language, DAX, Power Pivot, VBA macros, Excel Tables, PivotTables, data analysis, spreadsheet automation, RPA, robotic process automation, COM automation, Windows automation, AI Excel assistant, natural language Excel, Excel API, Excel COM interop

## License

MIT License - See [LICENSE](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)
