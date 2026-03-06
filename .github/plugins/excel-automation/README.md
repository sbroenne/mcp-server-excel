# Excel Automation Plugin

Automate Microsoft Excel on Windows through natural language. 25 tools with 225+ operations covering Power Query, DAX, PivotTables, Charts, VBA, Tables, Ranges, Slicers, and more — all via Excel's native COM API.

## Installation

```bash
# Using Copilot CLI (when available via awesome-copilot marketplace)
copilot plugin install excel-automation@awesome-copilot
```

## What's Included

### Agents

| Agent | Description |
|-------|-------------|
| `excel-automation` | Automate Microsoft Excel on Windows through natural language. Create workbooks, write data, build PivotTables, charts, Power Query, DAX measures, VBA macros, and more using the Excel MCP Server. |

### Skills

| Skill | Description | Installation |
|-------|-------------|-------------|
| `excel-mcp` | Comprehensive Excel MCP Server skill with 225+ operations, workflow guidance, tool selection, and reference documentation for Power Query, DAX, PivotTables, Charts, and more. | `npx skills add sbroenne/mcp-server-excel --skill excel-mcp` |

## Prerequisites

- **Windows** with Microsoft Excel 2016+ installed
- **.NET 10 SDK** or later
- **MCP Server**: `dotnet tool install --global Sbroenne.ExcelMcp.McpServer`

## MCP Server Configuration

After installing the MCP Server, add to your editor's MCP configuration:

```json
{
  "servers": {
    "excel-mcp": {
      "command": "mcp-excel"
    }
  }
}
```

## Capabilities

- **📁 File Operations** — Create, open, save, and close workbooks (including IRM/AIP-protected files)
- **📋 Ranges** — Read/write values, formulas, formatting, validation, and protection
- **🎨 Tables** — Create Excel Tables, filter, sort, manage columns
- **📊 PivotTables** — Create, configure fields, calculated items/fields
- **📈 Charts** — Create and configure charts with series, labels, and trendlines
- **🔄 Power Query** — Import, evaluate, and manage M code with auto-formatting
- **📊 Data Model/DAX** — Create measures, relationships, and manage Power Pivot
- **📝 VBA** — Create modules, run macros, version control
- **🔌 Connections** — Manage OLEDB/ODBC data connections
- **🎚️ Slicers** — Interactive filtering for PivotTables and Tables
- **🧮 Calculation Mode** — Control auto-recalc for bulk write performance

## Source

This plugin is maintained at [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel).

## License

MIT
