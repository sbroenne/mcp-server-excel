---
layout: default
title: "Excel MCP Server - AI-Powered Excel Automation"
description: "Control Microsoft Excel with natural language through AI assistants like GitHub Copilot and Claude. Automate Power Query, DAX, VBA, PivotTables, and more. One-click install for Visual Studio Code."
keywords: "Excel automation, MCP server, AI Excel, Power Query, DAX measures, VBA macros, GitHub Copilot Excel, Claude Excel, Excel CLI, M code, REPL"
canonical_url: "https://sbroenne.github.io/mcp-server-excel/"
---

<div class="hero">
  <div class="container">
    <div class="hero-content">
      <img src="{{ '/assets/images/icon.png' | relative_url }}" alt="Excel MCP Server Icon" class="hero-icon">
      <h1 class="hero-title">Excel MCP Server</h1>
      <p class="hero-subtitle">AI-Powered Excel Automation</p>
      <p class="hero-description">Control Microsoft Excel with Natural Language through AI assistants like GitHub Copilot and Claude. One-click install for Visual Studio Code.</p>
    </div>
  </div>
</div>

<div class="container content-section" markdown="1">

**Automate Excel with AI - A Model Context Protocol (MCP) server for comprehensive Excel automation through conversational AI.**

## 🤔 What is This?

**Use natural language OR command-line to automate complex Excel tasks - your choice.**

Stop manually clicking through Excel menus for repetitive tasks. Instead, describe what you want in plain English:

**Data Transformation & Analysis:**
- *"Optimize all my Power Queries in this workbook for better performance"*
- *"Create a PivotTable from SalesData table showing top 10 products by region with sum and average"*
- *"Build a DAX measure calculating year-over-year growth with proper time intelligence"*

**Formatting & Styling (No Programming Required):**
- *"Format the revenue columns as currency, make headers bold with blue background, and add borders to the table"*
- *"Apply conditional formatting to highlight values above $10,000 in red and below $5,000 in yellow"*
- *"Convert this data range to an Excel Table with style TableStyleMedium2, add auto-filters, and create a totals row"*

**Workflow Automation:**
- *"Find all cells containing 'Q1 2024' and replace with 'Q1 2025', then sort the table by Date descending"*
- *"Add data validation dropdowns to the Status column with options: Active, Pending, Completed"*
- *"Merge the header cells, center-align them, and auto-fit all column widths to content"*

The AI assistant analyzes your request, generates the proper Excel automation commands, and executes them **directly in your Excel application** - no formulas or programming knowledge required.

## 🚀 Visual Studio Code Quick Start (1 Minute)

<p>One-click setup with GitHub Copilot integration</p>
<p><a href="https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp" class="button-link">Install from Marketplace</a></p>

## What is This Project?

**ExcelMcp** is a comprehensive Excel automation toolkit with two interfaces:

1. **MCP Server**: Enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands
2. **CLI Tool**: Provides direct command-line control for scripting, RPA, and CI/CD workflows

Both share the same core functionality: automate Power Query, DAX measures, VBA macros, PivotTables, formatting, and data transformations. Choose MCP for AI-powered conversations or CLI for programmatic control.

**🛡️ 100% Safe - Uses Excel's Native COM API**

Unlike third-party libraries that manipulate `.xlsx` files directly (risking file corruption), ExcelMcp uses **Excel's official COM API**. This ensures:
- ✅ **Zero risk of document corruption** - Excel handles all file operations safely
- ✅ **Interactive development** - See changes in real-time, create → test → refine → iterate instantly
- ✅ **Comprehensive automation** - Currently supports 163 operations across 12 specialized tools covering Power Query, Data Model/DAX, VBA, PivotTables, Excel Tables, ranges, conditional formatting, and more

**💻 For Developers:** Think of Excel as an AI-powered REPL - write code (Power Query M, DAX, VBA), execute instantly, inspect results visually in the live workbook. No more blind editing of .xlsx files.

## Key Features

<div class="features-grid">
<div class="feature-card">
<h3>155 Operations</h3>
<p>11 specialized tools covering Power Query, DAX, VBA, PivotTables, ranges, conditional formatting, and more</p>
</div>

<div class="feature-card">
<h3>100% Safe</h3>
<p>Uses Excel's native COM API - zero corruption risk, full compatibility</p>
</div>

<div class="feature-card">
<h3>Interactive Development</h3>
<p>See changes in real-time - create, test, refine, and iterate instantly. Use Excel as a REPL.</p>
</div>

<div class="feature-card">
<h3>Natural Language Control</h3>
<p>Describe what you want, AI does the rest</p>
</div>

<div class="feature-card">
<h3>Comprehensive Automation</h3>
<p>155 operations covering Power Query, DAX, VBA, Tables, and more</p>
</div>

<div class="feature-card">
<h3>Dual Interface</h3>
<p>Choose MCP for AI assistants or CLI for scripts and RPA</p>
</div>
</div>

## Use Cases

### Data Transformation & Analysis
- Optimize Power Query M code for performance
- Create PivotTables from data with natural language
- Build DAX measures with AI guidance
- Transform and clean data automatically

👉 [See examples](#examples) of Power Query and DAX automation

### Formatting & Styling (No Programming)
- Format columns as currency, dates, percentages
- Apply conditional formatting with color rules
- Create Excel Tables with auto-filters
- Add data validation dropdowns

👉 [See examples](#examples) of formatting automation

### Workflow Automation
- Find and replace across ranges
- Sort and filter data programmatically
- Merge cells and auto-fit columns
- Manage Powerquery M Code, DAX statements and VBA macros with version control

## Installation

<div class="install-options">
<div class="install-option">
<h3>VS Code Extension <span class="badge">Recommended</span></h3>
<p>One-click setup with GitHub Copilot integration</p>
<p><a href="https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp" class="button-link">Install from Marketplace</a></p>
</div>

<div class="install-option">
<h3>MCP Server</h3>
<p>For Claude, ChatGPT, and other MCP clients</p>
<pre><code>dotnet tool install -g Sbroenne.ExcelMcp.McpServer</code></pre>
</div>

<div class="install-option">
<h3>CLI Tool</h3>
<p>For scripting, RPA, and CI/CD workflows</p>
<pre><code>dotnet tool install -g Sbroenne.ExcelMcp.CLI</code></pre>
</div>
</div>

📖 **[Complete Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md)**

## Documentation

- [Excel MCP Server source code on GitHub](https://github.com/sbroenne/mcp-server-excel)
- [How to contribute to Excel MCP Server development](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/CONTRIBUTING.md)
- [Complete API documentation and developer guides](https://github.com/sbroenne/mcp-server-excel/tree/main/docs)

## Examples

### AI Assistant Examples (MCP Server)

<div class="example-section">
<h4>🔍 Optimize Power Query</h4>
<p><strong>You:</strong> "This Power Query is taking 5 minutes to refresh. Can you optimize it?"</p>
<p>AI analyzes your M code, identifies inefficiencies, and applies best practices automatically.</p>
</div>

<div class="example-section">
<h4>📊 Create PivotTable</h4>
<p><strong>You:</strong> "Create a PivotTable from SalesData showing top 10 products by region"</p>
<p>AI creates the PivotTable with proper field configuration in seconds.</p>
</div>

<div class="example-section">
<h4>🎨 Format Data</h4>
<p><strong>You:</strong> "Format revenue as currency, make headers bold blue, and add borders"</p>
<p>AI applies all formatting directly to your Excel file.</p>
</div>

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

## Why Choose This Project?

<div class="features-grid">
<div class="feature-card">
<h3>🤖 MCP Server (AI-Powered)</h3>
<ul>
<li><strong>Natural Language Control:</strong> Describe tasks in plain English</li>
<li><strong>Safe:</strong> Official COM API - zero corruption risk</li>
<li><strong>Interactive:</strong> See changes in real-time in Excel</li>
<li><strong>Comprehensive:</strong> 155 operations across 11 tools</li>
<li><strong>Works with:</strong> GitHub Copilot, Claude, ChatGPT, and any MCP client</li>
</ul>
</div>

<div class="feature-card">
<h3>⚙️ CLI (Automation & Scripting)</h3>
<ul>
<li><strong>RPA Ready:</strong> Perfect for robotic process automation workflows</li>
<li><strong>CI/CD Integration:</strong> Automate Excel operations in build pipelines</li>
<li><strong>Scripting:</strong> PowerShell, Bash, Python integration</li>
<li><strong>Batch Processing:</strong> Process multiple workbooks programmatically</li>
<li><strong>No AI Required:</strong> Direct command-line control for scripts</li>
</ul>
</div>
</div>

<div class="callout">
<strong>Both Share the Same Core</strong>
Same 155 operations available via CLI and MCP. Consistent behavior across interfaces. Full .NET library available for custom integrations.
</div>

## Related Projects

- [Model Context Protocol specification and documentation](https://modelcontextprotocol.io/)
- [Official MCP server examples and implementations](https://github.com/modelcontextprotocol/servers)
- [Download Claude Desktop AI assistant with MCP support](https://claude.ai/download)
- [GitHub Copilot AI coding assistant](https://github.com/features/copilot)

## Keywords

<div class="keywords">
Excel automation, MCP server, Model Context Protocol, GitHub Copilot, Claude AI, Power Query, M language, DAX, Power Pivot, VBA macros, Excel Tables, PivotTables, data analysis, spreadsheet automation, RPA, robotic process automation, COM automation, Windows automation, AI Excel assistant, natural language Excel, Excel API, Excel COM interop, Excel CLI, command line Excel, Excel scripting, Excel batch processing, CI/CD Excel, DevOps Excel, Excel PowerShell, Excel automation tool, .NET Excel library, Excel NuGet package
</div>

</div>

<footer>
<div class="container">
<p><strong>Excel MCP Server</strong> - MIT License</p>
<p><a href="https://github.com/sbroenne/mcp-server-excel" title="View Excel MCP Server source code and documentation">GitHub Repository</a> | <a href="https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md" title="Installation instructions for Windows, VS Code, and CLI">Installation Guide</a> | <a href="https://github.com/sbroenne/mcp-server-excel/blob/main/docs/CONTRIBUTING.md" title="Learn how to contribute to Excel MCP Server">Contributing</a></p>
</div>
</footer>
