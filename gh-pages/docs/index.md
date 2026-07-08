---
template: home.html
title: Home
description: >-
  Control Microsoft Excel with natural language through AI assistants like
  GitHub Copilot and Claude. Automate Power Query, DAX, VBA, PivotTables and
  more — no Excel programming knowledge required.
keywords: "Excel automation, MCP server, AI Excel, Power Query, DAX measures, VBA macros, GitHub Copilot Excel, Claude Excel, Excel CLI, M code"
hide:
  - navigation
  - toc
---

**A Model Context Protocol (MCP) server and CLI for comprehensive Excel
automation through conversational AI.** Excel MCP Server lets AI assistants —
GitHub Copilot, Claude, ChatGPT and any MCP-compatible client — drive Microsoft
Excel with natural language: Power Query &amp; M, PowerPivot &amp; DAX, VBA
macros, PivotTables, charts, formatting and much more. No Excel programming
knowledge required.

!!! tip "100% safe — uses Excel's native COM API"
    Zero risk of file corruption. Unlike third-party libraries that manipulate
    `.xlsx` files directly, this project drives Excel's official COM API,
    ensuring complete safety and compatibility. Watch Excel update in real time
    as the AI works — just say *"Show me Excel while you work."*

<div class="mcp-video" markdown>
[![Watch the Excel MCP Server intro video](https://img.youtube.com/vi/B6eIQ5BIbNc/maxresdefault.jpg){ width="560" }](https://youtu.be/B6eIQ5BIbNc)

▶️ [Watch the intro video (1 min)](https://youtu.be/B6eIQ5BIbNc)
</div>

## Key features

<div class="grid cards" markdown>

-   :material-transit-connection-variant: __Power Query &amp; M code__

    ---

    Create, edit and optimize M code. Import from files, databases and APIs.
    Refresh queries and manage load destinations.

-   :material-calculator-variant: __Power Pivot &amp; DAX__

    ---

    Build Data Models, create DAX measures and manage table relationships.
    Full Power Pivot automation.

-   :material-chart-box: __PivotTables &amp; charts__

    ---

    Create PivotTables from ranges, tables or the Data Model. Build charts and
    PivotCharts with full formatting control.

-   :material-table: __Tables &amp; ranges__

    ---

    Read/write data, formulas and formatting. Filter, sort and validate. Manage
    Excel Tables with structured references.

-   :material-code-braces: __VBA macros__

    ---

    View, import, update and execute VBA code. Export modules for version
    control.

-   :material-file-table-box-multiple: __Worksheets &amp; connections__

    ---

    Manage sheets, named ranges and data connections. Copy and move sheets
    between workbooks.

-   :material-eye-outline: __Agent mode__

    ---

    Watch AI work in Excel in real time — side-by-side view, live status-bar
    feedback and smart window arrangement, like a pair programmer in a
    spreadsheet.

-   :material-test-tube: __LLM-tested quality__

    ---

    Tool behavior validated with real LLM workflows using
    [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering).
    We test that LLMs correctly understand and use our tools.

-   :fontawesome-brands-python: __Python in Excel__

    ---

    Write and run `=PY()` formulas that execute in Excel's cloud Python engine —
    process worksheet data with pandas, NumPy and more, from your AI assistant.

</div>

[See all 26 tools and 232 operations :material-arrow-right:](features.md){ .md-button }

## What can you do with it?

Ask your AI assistant to automate Excel tasks using natural language:

!!! example "📝 Create &amp; populate data"
    **You:** "Create a new Excel file with a table for tracking sales — include
    Date, Product, Quantity, Unit Price and Total with sample data and formulas."

    The AI creates the workbook, adds headers, enters sample data and builds
    formulas automatically.

!!! example "📊 PivotTables &amp; charts"
    **You:** "Create a PivotTable showing total sales by Product, then add a bar
    chart to visualize the results."

    The AI creates the PivotTable with proper field configuration and adds a
    linked chart.

!!! example "🔄 Power Query &amp; Data Model"
    **You:** "Use Power Query to import products.csv, load it to the Data Model,
    and create measures for Total Revenue and Average Rating."

    The AI imports the data, adds it to Power Pivot and creates DAX measures
    ready for analysis.

!!! example "🎨 Formatting &amp; tables"
    **You:** "Format the Price column as currency, highlight values over $500 in
    green, and convert this to an Excel Table."

    The AI applies number formats, conditional styling, auto-fit and structured
    table styling.

## CLI or MCP Server?

This package ships **both** a CLI and an MCP Server. They share the same core,
so every operation behaves identically — pick the entry point that fits your
workflow:

| Interface | Best for | Why |
|-----------|----------|-----|
| **CLI** (`excelcli`) | Coding agents (Copilot, Cursor, Windsurf) | **64% fewer tokens** — single tool, no large schemas. Better for cost-sensitive, high-throughput automation. |
| **MCP Server** | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery and a persistent connection. Better for interactive, exploratory workflows. |

[MCP Server docs](mcp-server.md){ .md-button } [CLI docs](cli.md){ .md-button }

## How it works

**Excel MCP Server uses Windows COM automation to control the actual Excel
application — not just `.xlsx` files.** The MCP Server and CLI are two equal,
first-class entry points. Each hosts its own service: the MCP Server runs it
**in-process** (direct calls, no pipe), while the CLI uses a **background
daemon** over a named pipe so sessions persist across `excelcli` invocations.

```text
┌──────────────────────┐        ┌──────────────────────┐
│  MCP Server          │        │  CLI (excelcli)      │
│  (AI assistants)     │        │  (coding agents)     │
└──────────┬───────────┘        └──────────┬───────────┘
           │ in-process                    │ named pipe →
           │ (direct calls)                │ background daemon
           ▼                               ▼
┌──────────────────────┐        ┌──────────────────────┐
│  ExcelMCP Service    │        │  ExcelMCP Service    │
│  (session mgmt)      │        │  (daemon; sessions   │
│                      │        │   persist across     │
│                      │        │   CLI invocations)   │
└──────────┬───────────┘        └──────────┬───────────┘
           ▼                               ▼
      Core Commands                   Core Commands
           ▼                               ▼
┌──────────────────────┐        ┌──────────────────────┐
│  Excel COM API       │        │  Excel COM API       │
│  (Excel.Application) │        │  (Excel.Application) │
└──────────────────────┘        └──────────────────────┘
```

Both entry points share the same Core Commands codebase, so every operation
behaves identically. They run as separate processes, each with its own service
and Excel instance, and do **not** share live sessions with each other.

## Documentation

<div class="grid cards" markdown>

-   :material-star-shooting: __[Feature reference](features.md)__

    All 26 tools and 232 operations, grouped by category.

-   :material-download: __[Installation guide](installation.md)__

    Setup for VS Code, Claude Desktop, other MCP clients and the CLI.

-   :material-connection: __[MCP Server](mcp-server.md)__

    Complete MCP tool reference and examples.

-   :material-console: __[CLI](cli.md)__

    Full CLI command reference and examples.

-   :material-robot-happy: __[Agent skills](skills.md)__

    Cross-platform AI guidance for 43+ agents.

-   :material-history: __[Changelog](changelog.md)__

    Release notes and version history.

</div>

## Related projects

Other projects by the author:

- [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering) — LLM-powered testing framework for AI agents.
- [Windows MCP Server](https://windowsmcpserver.dev/) — AI-powered Windows automation (mouse, keyboard, windows, screenshots).
- [OBS Studio MCP Server](https://github.com/sbroenne/mcp-server-obs) — AI-powered OBS Studio automation for recording and streaming.
- [RVToolsMerge](https://github.com/sbroenne/RvToolsMerge) — Merge and anonymize VMware RVTools exports.
- [Azure Retail Prices Exporter](https://github.com/sbroenne/azureretailprices-exporter) — Daily automated Azure pricing exports with FX rates.
- [AWS CUR Anonymize](https://github.com/sbroenne/aws-cur-anonymize) — Anonymize AWS Cost &amp; Usage Reports for secure sharing.
