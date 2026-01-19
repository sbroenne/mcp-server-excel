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
    </div>
  </div>
</div>

<div class="container content-section" markdown="1">
## 🤔 What is This?

**Automate Excel with AI - A Model Context Protocol (MCP) server for comprehensive Excel automation through conversational AI.**

<p>One-click setup with GitHub Copilot integration</p>
<p><a href="https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp" class="button-link">Install from Marketplace</a></p>

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
<p>Tool behavior validated with real AI agents using <a href="https://github.com/mykhaliev/agent-benchmark">agent-benchmark</a>. We test that LLMs correctly understand and use our tools.</p>
</div>
</div>

<p><a href="/features/">See all 21 tools and 187 operations →</a></p>

## What Can You Do With It?

Ask your AI assistant to automate Excel tasks using natural language:

<div class="example-section">
<h4>🔍 Power Query & M Code</h4>
<p><strong>You:</strong> "This Power Query is taking 5 minutes to refresh. Can you optimize it?"</p>
<p>AI analyzes your M code, identifies inefficiencies, and applies best practices automatically.</p>
</div>

<div class="example-section">
<h4>📊 PivotTables & Data Model</h4>
<p><strong>You:</strong> "Create a PivotTable from the Data Model showing top 10 products by region with a DAX measure for profit margin"</p>
<p>AI creates the PivotTable, adds fields, and builds the DAX measure in seconds.</p>
</div>

<div class="example-section">
<h4>📈 Charts & Visualization</h4>
<p><strong>You:</strong> "Create a column chart from SalesData and add a trend line"</p>
<p>AI builds the chart with proper formatting and data series configuration.</p>
</div>

<div class="example-section">
<h4>🖥️ VBA Macros</h4>
<p><strong>You:</strong> "Write a VBA macro that exports all sheets to PDF"</p>
<p>AI creates, imports, and can execute VBA code directly in your workbook.</p>
</div>

<div class="example-section">
<h4>🎨 Tables & Formatting</h4>
<p><strong>You:</strong> "Convert this range to an Excel Table, format revenue as currency, and add conditional formatting for values over 10000"</p>
<p>AI applies structured tables, number formats, and conditional rules.</p>
</div>

## CLI Tool (Optional)

For scripting, RPA workflows, and CI/CD pipelines — automate Excel without AI:

```bash
dotnet tool install -g Sbroenne.ExcelMcp.CLI
```

```bash
# Refresh Power Query
excel-cli pq refresh sales.xlsx --query SalesData

# Export VBA for version control  
excel-cli vba export macro-workbook.xlsm --module Module1 --output src/vba/

# Read range values
excel-cli range get-values report.xlsx --sheet Sheet1 --range A1:D10
```

The CLI provides 172 operations across 13 command groups, sharing the same Core engine as the MCP Server.


## Documentation

📖 **[Complete Feature Reference](/features/)** — All 21 tools and 187 operations

📥 **[Installation Guide](/installation/)** — Setup for VS Code, Claude Desktop, other MCP clients, and CLI

📋 **[Changelog](/changelog/)** — Release notes and version history

## More Information

- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel) — Source code, issues, and contributions
- [Contributing Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/CONTRIBUTING.md) — How to contribute

## Related Projects

Other projects by the author:

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
