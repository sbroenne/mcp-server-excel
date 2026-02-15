---
layout: default
title: "Automate Power Query, DAX & VBA with AI"
description: "Control Microsoft Excel with natural language through AI assistants like GitHub Copilot and Claude. Automate Power Query, DAX, VBA, PivotTables, and more. One-click install for Visual Studio Code."
keywords: "Excel automation, MCP server, AI Excel, Power Query, DAX measures, VBA macros, GitHub Copilot Excel, Claude Excel, Excel CLI, M code, REPL"
canonical_url: "https://excelmcpserver.dev/"
---

<div class="hero">
  <div class="container">
    <div class="hero-content">
      <img src="{{ '/assets/images/icon.png' | relative_url }}" alt="Excel MCP Server Icon" class="hero-icon">
      <h1 class="hero-title">Excel MCP Server</h1>
      <p class="hero-subtitle">Automate Excel with AI via GitHub Copilot, Claude, and other MCP clients — including Power Query, DAX, VBA, PowerPivot, and Tables.</p>
    </div>
  </div>
</div>

<div class="badges-section">
  <div class="container">
    <div class="hero-badges">
      <a href="https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp"><img src="https://img.shields.io/visual-studio-marketplace/i/sbroenne.excel-mcp?label=VS%20Code%20Installs" alt="VS Code Marketplace Installs"></a>
      <a href="https://github.com/sbroenne/mcp-server-excel"><img src="https://img.shields.io/github/stars/sbroenne/mcp-server-excel?style=flat&label=GitHub%20Stars" alt="GitHub Stars"></a>
      <a href="https://github.com/sbroenne/mcp-server-excel/releases"><img src="https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=GitHub%20Downloads" alt="GitHub Downloads"></a>
      <a href="https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer"><img src="https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.McpServer.svg?label=NuGet%20MCP%20Installs" alt="NuGet MCP Server Installs"></a>
      <a href="https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI"><img src="https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.CLI.svg?label=NuGet%20CLI%20Installs" alt="NuGet CLI Installs"></a>
    </div>
  </div>
</div>

<div class="container content-section" markdown="1">
## 🤔 What is This?

**Automate Excel with AI - A Model Context Protocol (MCP) server for comprehensive Excel automation through conversational AI.**

<div class="quick-install-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin: 24px 0;">
  <div style="text-align: center;">
    <p><strong>VS Code / GitHub Copilot</strong></p>
    <a href="https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp" class="button-link">Install Extension</a>
  </div>
  <div style="text-align: center;">
    <p><strong>Claude Desktop</strong></p>
    <a href="https://github.com/sbroenne/mcp-server-excel/releases/latest" class="button-link">One-Click Install</a>
  </div>
  <div style="text-align: center;">
    <p><strong>Cursor, Windsurf, etc.</strong></p>
    <a href="/installation/" class="button-link">Installation Guide</a>
  </div>
</div>

**MCP Server for Excel** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands, including Power Query & M, PowerPivot & DAX, VBA macros, PivotTables, Charts, formatting & much more – no Excel programming knowledge required.

It works with any MCP-compatible AI assistant like GitHub Copilot, Claude Desktop, Cursor, Windsurf, etc.

**🛡️ 100% Safe - Uses Excel's Native COM API** - Zero risk of file corruption. Unlike third-party libraries that manipulate `.xlsx` files directly, this project uses Excel's official API ensuring complete safety and compatibility.

**💡 Interactive Development** - Watch Excel update in real-time as AI works. Say "Show me Excel while you work" and see every change live - create a query, watch it populate, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing.

<div class="video-preview">
  <a href="https://youtu.be/B6eIQ5BIbNc" target="_blank" title="Watch Excel MCP Server intro video">
    <img src="https://img.youtube.com/vi/B6eIQ5BIbNc/maxresdefault.jpg" alt="Excel MCP Server intro video thumbnail" style="max-width: 560px; width: 100%; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
    <p style="text-align: center; margin-top: 8px;">▶️ Watch the intro video (1 min)</p>
  </a>
</div>

## Key Features

<div class="features-grid">
<div class="feature-card">
<h3>Power Query & M Code</h3>
<p>Create, edit, and optimize M code. Import from files, databases, APIs. Refresh queries and manage load destinations.</p>
</div>

<div class="feature-card">
<h3>Power Pivot & DAX</h3>
<p>Build Data Models, create DAX measures, manage table relationships. Full Power Pivot automation.</p>
</div>

<div class="feature-card">
<h3>PivotTables & Charts</h3>
<p>Create PivotTables from ranges, tables, or Data Model. Build charts and PivotCharts with full formatting control.</p>
</div>

<div class="feature-card">
<h3>Tables & Ranges</h3>
<p>Read/write data, formulas, formatting. Filter, sort, validate. Manage Excel Tables with structured references.</p>
</div>

<div class="feature-card">
<h3>VBA Macros</h3>
<p>View, import, update, and execute VBA code. Export modules for version control.</p>
</div>

<div class="feature-card">
<h3>Worksheets & Connections</h3>
<p>Manage sheets, named ranges, data connections. Copy/move sheets between workbooks.</p>
</div>

<div class="feature-card">
<h3>🧪 LLM-Tested Quality</h3>
<p>Tool behavior validated with real LLM workflows using <a href="https://github.com/sbroenne/pytest-aitest">pytest-aitest</a>. We test that LLMs correctly understand and use our tools.</p>
</div>
</div>

<p><a href="/features/">See all 23 tools and 214 operations →</a></p>

## What Can You Do With It?

Ask your AI assistant to automate Excel tasks using natural language:

<div class="example-section">
<h4>📝 Create & Populate Data</h4>
<p><strong>You:</strong> "Create a new Excel file with a table for tracking sales - include Date, Product, Quantity, Unit Price, and Total with sample data and formulas."</p>
<p>AI creates the workbook, adds headers, enters sample data, and builds formulas automatically.</p>
</div>

<div class="example-section">
<h4>📊 PivotTables & Charts</h4>
<p><strong>You:</strong> "Create a PivotTable showing total sales by Product, then add a bar chart to visualize the results."</p>
<p>AI creates the PivotTable with proper field configuration and adds a linked chart.</p>
</div>

<div class="example-section">
<h4>🔄 Power Query & Data Model</h4>
<p><strong>You:</strong> "Use Power Query to import products.csv, load it to the Data Model, and create measures for Total Revenue and Average Rating."</p>
<p>AI imports the data, adds it to Power Pivot, and creates DAX measures ready for analysis.</p>
</div>

<div class="example-section">
<h4>🎛️ Interactive Filtering</h4>
<p><strong>You:</strong> "Create a slicer for the Region field so I can filter the PivotTable interactively."</p>
<p>AI adds a slicer connected to your PivotTable for point-and-click filtering.</p>
</div>

<div class="example-section">
<h4>🎨 Formatting & Tables</h4>
<p><strong>You:</strong> "Format the Price column as currency, highlight values over $500 in green, and convert this to an Excel Table."</p>
<p>AI applies number formats, conditional formatting, and structured table styling.</p>
</div>

## CLI vs MCP Server

This package provides both **CLI** and **MCP Server** interfaces. Choose based on your use case:

| Interface | Best For | Why |
|-----------|----------|-----|
| **CLI** (`excelcli`) | Coding agents (Copilot, Cursor, Windsurf) | **64% fewer tokens** - single tool, no large schemas. Better for cost-sensitive, high-throughput automation. |
| **MCP Server** | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent connection. Better for interactive, exploratory workflows. |

**Manual Installation:**
```powershell
# Step 1: Install MCP Server and CLI
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
dotnet tool install --global Sbroenne.ExcelMcp.CLI

# Step 2: Auto-configure all your coding agents
npx add-mcp "mcp-excel" --name excel-mcp

# Optional: Install agent skills for better AI guidance
npx skills add sbroenne/mcp-server-excel
```

> 💡 **Skills provide AI guidance** - The CLI skill is highly recommended (agents don't work perfectly with CLI without it). The MCP skill adds workflow best practices and reduces token usage.


## ⚙️ How It Works - Unified Service Architecture

**ExcelMcp uses Windows COM automation to control the actual Excel application (not just .xlsx files).**

Both the **MCP Server** and **CLI** communicate with a shared **ExcelMCP Service** that manages Excel sessions:

```
┌─────────────────────┐     ┌─────────────────────┐
│   MCP Server        │     │   CLI (excelcli)    │
│  (AI assistants)    │     │  (coding agents)    │
└─────────┬───────────┘     └─────────┬───────────┘
          │                           │
          └──────────┬────────────────┘
                     ▼
          ┌─────────────────────────┐
          │   ExcelMCP Service      │
          │  (shared session mgmt)  │
          └─────────┬───────────────┘
                    ▼
          ┌─────────────────────────┐
          │   Excel COM API         │
          │  (Excel.Application)    │
          └─────────────────────────┘
```

**Key Benefits:**
- ✅ **Shared Sessions** - CLI and MCP Server can access the same open workbooks
- ✅ **Single Excel Instance** - No duplicate Excel processes or file locks
- ✅ **System Tray UI** - Monitor active sessions via the ExcelMCP tray icon

**💡 Tip: Watch Excel While AI Works**
By default, Excel runs hidden for faster automation. To see changes in real-time, just ask:
- *"Show me Excel while you work"*
- *"Let me watch what you're doing"*
- *"Open Excel so I can see the changes"*

The AI will display the Excel window so you can watch every operation happen live!


## Documentation

📖 **[Complete Feature Reference](/features/)** — All 23 tools and 214 operations

📥 **[Installation Guide](/installation/)** — Setup for VS Code, Claude Desktop, other MCP clients, and CLI

📖 **[MCP Server Documentation](/mcp-server/)** — Complete MCP tool reference and examples

📖 **[CLI Documentation](/cli/)** — Full CLI command reference and examples

🤖 **[Agent Skills](/skills/)** — Cross-platform AI guidance for 43+ agents (auto-installed by VS Code extension)

📋 **[Changelog](/changelog/)** — Release notes and version history

## More Information

- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel) — Source code, issues, and contributions
- [Contributing Guide](/contributing/) — How to contribute

## Related Projects

Other projects by the author:

- [pytest-aitest](https://github.com/sbroenne/pytest-aitest) — LLM-powered testing framework for AI agents — validate that LLMs correctly understand and use your tools
- [Windows MCP Server](https://windowsmcpserver.dev/) — AI-powered Windows automation via GitHub Copilot, Claude, and other MCP clients — including mouse, keyboard, windows, and screenshots
- [OBS Studio MCP Server](https://github.com/sbroenne/mcp-server-obs) — AI-powered OBS Studio automation for recording, streaming, and window capture
- [HeyGen MCP Server](https://github.com/sbroenne/heygen-mcp) — MCP server for HeyGen AI video generation
- [RVToolsMerge](https://github.com/sbroenne/RvToolsMerge) — Merge and anonymize VMware RVTools exports.
- [Azure Retail Prices Exporter](https://github.com/sbroenne/azureretailprices-exporter) — Daily automated Azure pricing data exports with FX rates
- [AWS CUR Anonymize](https://github.com/sbroenne/aws-cur-anonymize) — Anonymize AWS Cost & Usage Reports for secure sharing
  
<footer>
<div class="container">
<p><strong>Excel MCP Server</strong> — MIT License — © 2024-2025</p>
</div>
</footer>
