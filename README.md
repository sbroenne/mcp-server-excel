# ExcelMcp - MCP Server for Microsoft Excel

[![VS Code Marketplace Installs](https://vsmarketplacebadges.dev/installs-short/sbroenne.excel-mcp.svg?label=VS%20Code%20Installs)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=GitHub%20Downloads)](https://github.com/sbroenne/mcp-server-excel/releases)

[![CI Gate](https://github.com/sbroenne/mcp-server-excel/actions/workflows/ci.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-10-blue.svg)](https://dotnet.microsoft.com/download/dotnet/10.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

[**Website**](https://excelmcpserver.dev/) ·
[**Installation**](https://excelmcpserver.dev/installation/) ·
[**Features**](https://excelmcpserver.dev/features/) ·
[**Troubleshooting**](https://excelmcpserver.dev/troubleshooting/) ·
[**1-minute demo**](https://youtu.be/B6eIQ5BIbNc)

**Automate real Microsoft Excel with AI.** Excel MCP Server lets GitHub Copilot,
Claude, ChatGPT, and other agents control Excel through natural-language
requests—using either MCP or a token-efficient CLI.

Unlike file-parser tools, ExcelMcp drives the **actual Excel application** through
its official COM API. It can refresh Power Query, recalculate formulas, evaluate
DAX, run VBA and Python `=PY()`, and preserve PivotTables, charts, macros, the
Data Model, and workbook formatting.

**31 tools with 329 operations** cover end-to-end Excel automation.

> [!IMPORTANT]
> Requires **Windows**, **Microsoft Excel 2016 or later**, and an interactive
> desktop. It is not intended for Linux, macOS, or server-side batch processing.

## 🚀 Get Started

| Use case | Recommended path |
|---|---|
| **VS Code** | [Install the extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) |
| **Claude Desktop or another MCP client** | [Install the MCP Server](https://excelmcpserver.dev/installation-mcp-server/) |
| **Coding agents and scripts** | [Install the CLI](https://excelmcpserver.dev/installation-cli/) |
| **Not sure which to choose?** | [Read the installation overview](https://excelmcpserver.dev/installation/) |

Close open Excel workbooks before starting; ExcelMcp requires exclusive access
while automating them.

## What You Can Automate

- **[Data & analytics](https://excelmcpserver.dev/features/data-analytics/):**
  Power Query, DAX, Power Pivot, Excel Tables, PivotTables, and data connections.
- **[Cells & workbooks](https://excelmcpserver.dev/features/cells-workbooks/):**
  ranges, formulas, formatting, worksheets, files, calculation, and named ranges.
- **[Charts & visuals](https://excelmcpserver.dev/features/charts-visuals/):**
  charts, slicers, conditional formatting, screenshots, drawings, and sparklines.
- **[Automation & advanced](https://excelmcpserver.dev/features/automation-advanced/):**
  VBA, Python in Excel, Goal Seek, scenarios, data tables, windows, and XML Maps.

Explore the [complete reference for all 329 operations](https://excelmcpserver.dev/features/).

## See It in Action

[![A sales table, regional summary, and chart created in the real Excel application by Excel MCP Server](https://excelmcpserver.dev/assets/images/excel-demo-table-chart.png)](https://excelmcpserver.dev/use-cases/)

Ask in plain language:

- *"Import products.csv with Power Query and load it to the Data Model."*
- *"Create a PivotTable showing revenue by region, then add a column chart."*
- *"Use Goal Seek to find the price that produces $100,000 profit."*
- *"Show me Excel while you work."*

Because Excel performs the work, you can inspect results live and continue editing
the workbook normally.

[See more examples and use cases](https://excelmcpserver.dev/use-cases/).

## MCP Server or CLI?

Both entry points expose the same Core commands and behavior.

| Interface | Best for | Why |
|---|---|---|
| **MCP Server** | Conversational assistants and exploratory work | Rich schemas, tool discovery, and persistent sessions |
| **CLI (`excelcli`)** | Coding agents, automation, and scripts | One compact tool surface and substantially lower token usage |

The MCP Server calls the ExcelMcp service in-process. The CLI uses a background
daemon so workbook sessions persist across commands.

```text
AI assistant or script
        │
   MCP Server / CLI
        │
 ExcelMcp Core commands
        │
 Real Excel COM API
```

[Read the architecture](docs/ARCHITECTURE.md) or browse the
[MCP Server](https://excelmcpserver.dev/mcp-server/) and
[CLI](https://excelmcpserver.dev/cli/) guides.

## ⭐ GitHub Star History

[![GitHub stars over time for ExcelMcp](https://excelmcpserver.dev/assets/images/star-history.svg)](https://github.com/sbroenne/mcp-server-excel/stargazers)

## 📋 Additional Information

[Documentation](https://excelmcpserver.dev/) ·
[Changelog](https://excelmcpserver.dev/changelog/) ·
[Contributing](https://excelmcpserver.dev/contributing/) ·
[Security](https://excelmcpserver.dev/security/) ·
[Privacy](https://excelmcpserver.dev/privacy/)

**License:** MIT License - see [LICENSE](LICENSE) file

## Related Projects

Other projects by the author:

- [PowerPoint MCP Server](https://powerpointmcpserver.dev/) — AI-powered PowerPoint automation via MCP, the sister project to this one
- [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering) — LLM-powered testing framework for AI agents
- [Windows MCP Server](https://windowsmcpserver.dev/) — AI-powered Windows automation via MCP
- [OBS Studio MCP Server](https://github.com/sbroenne/mcp-server-obs) — AI-powered OBS Studio automation
