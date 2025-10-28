# ExcelMcp - MCP Server for Excel

[![VS Code Marketplace](https://img.shields.io/visual-studio-marketplace/v/sbroenne.excelmcp?label=VS%20Code%20Marketplace)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excelmcp)
[![Installs](https://img.shields.io/visual-studio-marketplace/i/sbroenne.excelmcp)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excelmcp)
[![GitHub](https://img.shields.io/badge/GitHub-sbroenne%2Fmcp--server--excel-blue)](https://github.com/sbroenne/mcp-server-excel)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Excel automation MCP server extension for Visual Studio Code**

This extension enables AI assistants like GitHub Copilot to interact with Microsoft Excel through the ExcelMcp MCP server. Automate Power Query M code, DAX measures, VBA macros, Excel Tables, ranges, worksheets, and data connections using natural language.

## Installation

### Option 1: VS Code Marketplace (Recommended)

1. Open VS Code
2. Go to Extensions (`Ctrl+Shift+X`)
3. Search for "ExcelMcp"
4. Click **Install**

### Option 2: Manual VSIX Install

1. Download the `.vsix` file from [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases)
2. In VS Code: `Ctrl+Shift+P` → "Install from VSIX"
3. Select the downloaded file

## Features

- 🤖 **AI-Powered Excel Automation** - Control Excel through GitHub Copilot and other AI assistants
- 📊 **Power Query Management** - Create, view, update, and refactor M code
- 📈 **Data Model & DAX** - Manage measures, relationships, and calculated columns
- 📋 **Excel Tables** - 22 operations for table lifecycle, columns, filters, and sorts
- 🔧 **VBA Macros** - List, export, import, and execute VBA code
- 🔌 **Data Connections** - Manage OLEDB, ODBC, Text, and Web connections
- 📐 **Ranges & Worksheets** - 30+ operations for data manipulation

## Requirements

- **Windows OS** - Excel COM automation requires Windows
- **Microsoft Excel 2016+** - Must be installed on your system
- **.NET 8 Runtime** - **Automatically installed** by the extension via the .NET Install Tool

### Prerequisites

Only Excel needs to be installed - the extension handles .NET automatically!

## Quick Start

1. **Install this extension** from the VS Code Marketplace or VSIX file
2. **The extension automatically**:
   - Installs .NET 8 runtime (via .NET Install Tool)
   - Installs the ExcelMcp MCP server tool
   - Configures the MCP server for AI assistants
3. **Start using** - Ask GitHub Copilot to help with Excel tasks:
   - "List all Power Query queries in workbook.xlsx"
   - "Export all DAX measures to .dax files"
   - "Create a new Excel table from range A1:D100"

**No manual setup required!** The extension handles everything automatically.

## What is MCP?

The Model Context Protocol (MCP) is an open standard that enables AI assistants to interact with external tools and data sources. This extension registers the ExcelMcp MCP server with VS Code, making Excel automation available to AI coding assistants.

## Available Tools

The ExcelMcp MCP server provides **10 specialized tools**:

1. **excel_powerquery** - Power Query M code (11 actions)
2. **excel_datamodel** - DAX measures & relationships (20 actions)
3. **table** - Excel Tables/ListObjects (22 actions)
4. **excel_range** - Range operations (30+ actions)
5. **excel_vba** - VBA macros (7 actions)
6. **excel_connection** - Data connections (11 actions)
7. **excel_worksheet** - Worksheet lifecycle (5 actions)
8. **excel_parameter** - Named ranges (6 actions)
9. **excel_file** - File creation (1 action)
10. **excel_version** - Update checking (1 action)

## How It Works

This extension uses the **NuGet MCP approach**:

- The extension registers the MCP server with VS Code
- When an AI assistant needs Excel automation, VS Code runs: `dnx Sbroenne.ExcelMcp.McpServer --yes`
- The `dnx` command automatically downloads the latest version from NuGet
- The MCP server communicates with Excel via COM automation

## Configuration

The extension works out-of-the-box. No manual configuration needed!

If you want to verify or customize the MCP server configuration, you can check VS Code's MCP settings.

## Documentation

- **[Main Repository](https://github.com/sbroenne/mcp-server-excel)** - Complete documentation
- **[MCP Server Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.McpServer/README.md)** - Detailed tool reference
- **[CLI Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/CLI.md)** - Command-line usage

## Example Use Cases

### Power Query Development
```
You: "This Power Query is slow. Can you refactor it?"
Copilot: [Analyzes M code → Suggests optimizations → Updates query]
```

### Data Model Management
```
You: "Create a new DAX measure for total sales"
Copilot: [Creates measure with proper format → Exports to .dax file]
```

### VBA Enhancement
```
You: "Add error handling to this VBA module"
Copilot: [Exports VBA → Enhances code → Updates module]
```

## Support

- 🐛 **Report Issues**: [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- 💬 **Discussions**: [GitHub Discussions](https://github.com/sbroenne/mcp-server-excel/discussions)
- 📖 **Documentation**: [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)

## License

MIT License - see [LICENSE](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)

## Acknowledgments

- Microsoft Excel Team - For comprehensive COM automation APIs
- Model Context Protocol community - For the AI integration standard
- Built with GitHub Copilot AI assistance
