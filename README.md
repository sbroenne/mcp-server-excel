# ExcelMcp - MCP Server for Microsoft Excel

[![Build MCP Server](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml)
[![Build CLI](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total)](https://github.com/sbroenne/mcp-server-excel/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-10.0-blue.svg)](https://dotnet.microsoft.com/download/dotnet/10.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

**A Model Context Protocol (MCP) server that gives AI assistants full control over Microsoft Excel through native COM automation.**

Control Power Query M code, Data Models with DAX measures, VBA macros, Excel Tables, connections, ranges, and worksheets through conversational AI. Also includes a CLI for direct human automation.

## 🎯 What You Can Do

**Power Query & M Code:**
- Create, read, update, delete Power Query transformations
- Export/import M code for version control
- Manage query load destinations (worksheet/data model/connection-only)
- Set privacy levels for data source combinations

**Data Model & DAX:**
- Create/update/delete DAX measures with format types (Currency, Percentage, Decimal, General)
- Manage table relationships (create, toggle active/inactive, delete)
- Discover model structure (tables, columns, measures, relationships)
- Export measures to .dax files for Git workflows

**Excel Tables (ListObjects):**
- 22 operations: create, resize, rename, delete, style
- Column management: add, remove, rename columns
- Data operations: append rows, apply filters (criteria/values), sort (single/multi-column)
- Advanced features: structured references, totals row, Data Model integration

**VBA Macros:**
- List, view, export, import, update VBA modules
- Execute macros with parameters
- Version control VBA code through file exports

**Ranges & Worksheets:**
- 30+ range operations: get/set values/formulas, clear, copy, insert/delete, find/replace, sort
- Manage hyperlinks and range properties
- Worksheet lifecycle: create, rename, copy, delete

**Data Connections:**
- Manage OLEDB, ODBC, Text, Web connections
- Update connection strings and properties
- Test connections and troubleshoot issues

## 🔧 How It Works - COM Interop Architecture

**ExcelMcp uses Windows COM automation to control the actual Excel application (not just .xlsx files).**

**✅ Benefits:**
- **Full Excel Feature Access** - Power Query engine, VBA runtime, Data Model, calculation engine, charts, pivot tables
- **True Compatibility** - Works exactly like Excel UI, no feature limitations
- **Live Data Refresh** - Can refresh Power Query, connections, Data Model in real workbooks
- **REPL Development** - Interactive development with immediate Excel feedback
- **No File Format Restrictions** - Handles .xlsx, .xlsm, .xlsb, legacy formats

**⚠️ Requirements:**
- **Windows Only** - COM interop is Windows-specific technology
- **Excel Installation Required** - Must have Microsoft Excel installed (2016 or later)
- **Desktop Automation** - Controls actual Excel process (not suitable for server-side processing)

**💡 When to Use:**
- Excel development and automation workflows
- Power Query/VBA/Data Model management
- Interactive Excel operations with AI assistance
- Version control integration for Excel code artifacts

**❌ When NOT to Use:**
- Server-side data processing (use file-based libraries instead)
- Linux/macOS environments (Excel not available)
- High-volume batch processing (consider Excel-free alternatives)

## 🚀 Quick Start

**Requirements:** Windows OS + Microsoft Excel 2016+ + .NET SDK 10

### 1. Install .NET 10 SDK - required for the "dnx" command

```powershell
winget install Microsoft.DotNet.SDK.10
```

### 2. Configure Your AI Assistant

**For GitHub Copilot** - Create `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "excel": {
      "command": "dnx",
      "args": ["Sbroenne.ExcelMcp.McpServer", "--yes"]
    }
  }
}
```

**For Claude Desktop** - Add to your MCP configuration:

```json
{
  "mcpServers": {
    "excel": {
      "command": "dnx",
      "args": ["Sbroenne.ExcelMcp.McpServer", "--yes"]
    }
  }
}
```

### 3. Verify Setup

Ask your AI assistant:
```
List all available Excel MCP tools
```

You should see 10 Excel tools: `excel_file`, `excel_powerquery`, `excel_connection`, `excel_worksheet`, `excel_range`, `excel_parameter`, `excel_vba`, `excel_datamodel`, `table`, `excel_version`.

**That's it!** The `dnx` command automatically downloads and runs the latest MCP server version.

> 💡 **First-time Setup Helper:** Download [excel-powerquery-vba-copilot-instructions.md](https://raw.githubusercontent.com/sbroenne/mcp-server-excel/main/instructions/excel-powerquery-vba-copilot-instructions.md) and save to `YourProject/.github/` for AI assistant guidance, or ask GitHub Copilot to set up everything automatically.

## 🔟 MCP Tools Overview

**10 specialized tools for comprehensive Excel automation:**

1. **excel_powerquery** (11 actions) - Power Query M code: create, view, import, export, update, delete, manage load destinations, privacy levels
2. **excel_datamodel** (15 actions) - Data Model & DAX: CRUD measures/relationships, discover structure, export to .dax files *[Phase 2: Full CRUD]*
3. **table** (22 actions) - Excel Tables: lifecycle, columns, filters, sorts, structured references, totals, Data Model integration *[Phase 2: Advanced]*
4. **excel_range** (30+ actions) - Ranges: get/set values/formulas, clear, copy, insert/delete, find/replace, sort, hyperlinks
5. **excel_vba** (7 actions) - VBA: list, view, export, import, update, run, delete modules
6. **excel_connection** (11 actions) - Connections: OLEDB/ODBC/Text/Web management, properties, refresh, test
7. **excel_worksheet** (5 actions) - Worksheets: list, create, rename, copy, delete
8. **excel_parameter** (6 actions) - Named ranges: list, get, set, create, delete, update
9. **excel_file** (1 action) - File creation: create empty .xlsx/.xlsm workbooks
10. **excel_version** (1 action) - Update checking from NuGet.org

> 📚 **[Complete MCP Server Guide →](src/ExcelMcp.McpServer/README.md)** - Detailed tool documentation and examples

---

## 📋 Additional Information

### CLI for Direct Automation

ExcelMcp also provides a command-line interface for script-based Excel automation (no AI required):

- **[ExcelMcp.CLI Guide](docs/CLI.md)** - Complete CLI documentation
- **[Command Reference](docs/COMMANDS.md)** - All 50+ CLI commands
- **Use Cases:** CI/CD pipelines, batch processing, scheduled tasks, PowerShell scripts

### Documentation

| Document | Description |
|----------|-------------|
| **[MCP Server Guide](src/ExcelMcp.McpServer/README.md)** | MCP setup, AI integration, examples |
| **[MCP Registry Publishing](docs/MCP_REGISTRY_PUBLISHING.md)** | How the server is published to MCP Registry |
| **[CLI Guide](docs/CLI.md)** | Command-line interface for automation |
| **[Command Reference](docs/COMMANDS.md)** | All 50+ CLI commands |
| **[Installation Guide](docs/INSTALLATION.md)** | Building from source |
| **[Development Workflow](docs/DEVELOPMENT.md)** | Contributing guidelines |
| **[Copilot Instructions](instructions/excel-powerquery-vba-copilot-instructions.md)** | AI assistant setup guide |

### Project Information

**License:** MIT License - see [LICENSE](LICENSE) file

**Contributing:** See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for guidelines

**Built With:** This entire project was developed using GitHub Copilot AI assistance

**Acknowledgments:**
- Microsoft Excel Team - For comprehensive COM automation APIs
- Model Context Protocol community - For the AI integration standard
- Open Source Community - For inspiration and best practices

---

### SEO & Discovery

`MCP Server` • `Model Context Protocol` • `Excel Automation` • `GitHub Copilot` • `AI Excel` • `Power Query` • `DAX Measures` • `Data Model` • `VBA Macros` • `Excel Tables` • `COM Interop` • `Windows Excel` • `Excel Development`
