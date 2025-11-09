# Excel MCP Server - AI-Powered Excel Automation

**Control Microsoft Excel with Natural Language through AI assistants like GitHub Copilot and Claude.**

## What is This Project?

**ExcelMcp** is a comprehensive Excel automation toolkit with two interfaces:

1. **MCP Server**: Enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands
2. **CLI Tool**: Provides direct command-line control for scripting, RPA, and CI/CD workflows

Both share the same core functionality: automate Power Query, DAX measures, VBA macros, PivotTables, formatting, and data transformations. Choose MCP for AI-powered conversations or CLI for programmatic control - no Excel programming knowledge required.

**💡 Interactive Development:** Unlike file-based tools, ExcelMcp lets you see results instantly in Excel - create → test → refine → iterate in real-time.

**💻 For Developers:** Think of Excel as an AI-powered REPL - write code (Power Query M, DAX, VBA), execute instantly, inspect results visually in the live workbook.

## Key Features

- **165 Operations** across 11 specialized tools
- **100% Safe** - Uses Excel's native COM API (zero corruption risk)
- **Interactive Development** - See changes in real-time, create → test → refine → iterate instantly
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

### Option 1: VS Code + GitHub Copilot (Recommended for AI Assistants)
Install the [Excel MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) for one-click setup with natural language control.

### Option 2: MCP Server (Any MCP Client - Claude, ChatGPT, etc.)
```bash
# Install globally as a .NET tool
dotnet tool install -g Sbroenne.ExcelMcp.McpServer

# Configure in your MCP client (e.g., Claude Desktop)
{
  "mcpServers": {
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "excel-mcp"]
    }
  }
}
```

### Option 3: CLI for Scripting & RPA (No AI Required)
```bash
# Install the CLI for automation scripts, CI/CD, and RPA workflows
dotnet tool install -g Sbroenne.ExcelMcp.CLI

# Use directly in scripts
excel-mcp pq-list --file "workbook.xlsx"
excel-mcp sheet-create --file "workbook.xlsx" --name "NewSheet"
excel-mcp range-set-values --file "workbook.xlsx" --sheet "Sheet1" --range "A1:C3" --values "[[1,2,3],[4,5,6],[7,8,9]]"
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

### AI Assistant Examples (MCP Server)

**Example 1: Optimize Power Query**
```
You: "This Power Query is taking 5 minutes to refresh. Can you optimize it?"

AI analyzes your M code, identifies inefficiencies, and applies best practices automatically.
```

**Example 2: Create PivotTable**
```
You: "Create a PivotTable from SalesData showing top 10 products by region"

AI creates the PivotTable with proper field configuration in seconds.
```

**Example 3: Format Data**
```
You: "Format revenue as currency, make headers bold blue, and add borders"

AI applies all formatting directly to your Excel file.
```

### CLI Examples (Scripting & RPA)

**Automate Power Query Refresh**
```bash
# Refresh all queries in a workbook (CI/CD pipeline)
excel-mcp pq-refresh --file "sales-report.xlsx" --query "SalesData"
```

**Batch Update Worksheets**
```powershell
# Process multiple workbooks in PowerShell
Get-ChildItem *.xlsx | ForEach-Object {
    excel-mcp sheet-create --file $_.Name --name "Summary"
    excel-mcp range-set-values --file $_.Name --sheet "Summary" --range "A1" --values "[[\"Generated: $(Get-Date)\"]]"
}
```

**Export VBA for Version Control**
```bash
# Export all VBA modules to Git repository
excel-mcp vba-export --file "macro-workbook.xlsm" --module "Module1" --output "src/vba/Module1.bas"
```

**Data Transformation Pipeline**
```bash
# Automated data processing pipeline
excel-mcp file-create --file "output.xlsx"
excel-mcp sheet-create --file "output.xlsx" --name "Processed"
excel-mcp range-set-values --file "output.xlsx" --sheet "Processed" --range "A1:C100" --values-from-json data.json
excel-mcp table-create --file "output.xlsx" --sheet "Processed" --range "A1:C100" --name "ProcessedData"
```

## Why Choose This Project?

### MCP Server (AI-Powered)
- **Natural Language Control**: Describe tasks in plain English
- **Safe**: Official COM API - zero corruption risk
- **Interactive**: See changes in real-time in Excel
- **Comprehensive**: 165 operations across 11 tools
- **Works with**: GitHub Copilot, Claude, ChatGPT, and any MCP client

### CLI (Automation & Scripting)
- **RPA Ready**: Perfect for robotic process automation workflows
- **CI/CD Integration**: Automate Excel operations in build pipelines
- **Scripting**: PowerShell, Bash, Python integration
- **Batch Processing**: Process multiple workbooks programmatically
- **No AI Required**: Direct command-line control for scripts

### Both Share the Same Core
- Same 165 operations available via CLI and MCP
- Consistent behavior across interfaces
- Full .NET library available for custom integrations

## Related Projects

- [Model Context Protocol](https://modelcontextprotocol.io/) - The protocol specification
- [MCP Servers](https://github.com/modelcontextprotocol/servers) - Official MCP server examples
- [Claude Desktop](https://claude.ai/download) - AI assistant with MCP support
- [GitHub Copilot](https://github.com/features/copilot) - AI coding assistant

## Keywords

Excel automation, MCP server, Model Context Protocol, GitHub Copilot, Claude AI, Power Query, M language, DAX, Power Pivot, VBA macros, Excel Tables, PivotTables, data analysis, spreadsheet automation, RPA, robotic process automation, COM automation, Windows automation, AI Excel assistant, natural language Excel, Excel API, Excel COM interop, Excel CLI, command line Excel, Excel scripting, Excel batch processing, CI/CD Excel, DevOps Excel, Excel PowerShell, Excel automation tool, .NET Excel library, Excel NuGet package

## License

MIT License - See [LICENSE](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)
