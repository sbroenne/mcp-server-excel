---
google-site-verification: 7BSNb2Q6rNfwdJCSGEUs9_2o3NOK09tBy4svR9A1bUg
layout: default
---

<link rel="stylesheet" href="{{ '/assets/css/style.css' | relative_url }}">
<link rel="icon" type="image/png" href="{{ '/assets/images/icon.png' | relative_url }}">

<div class="hero">
  <div class="container">
    <img src="{{ '/assets/images/icon.png' | relative_url }}" alt="Excel MCP Server" class="hero-icon">
    <h1>Excel MCP Server</h1>
    <p class="subtitle">Control Microsoft Excel with Natural Language through AI assistants like GitHub Copilot and Claude</p>
  </div>
</div>

<div class="container content-section">

## What is This Project?

**ExcelMcp** is a comprehensive Excel automation toolkit with two interfaces:

1. **MCP Server**: Enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands
2. **CLI Tool**: Provides direct command-line control for scripting, RPA, and CI/CD workflows

Both share the same core functionality: automate Power Query, DAX measures, VBA macros, PivotTables, formatting, and data transformations. Choose MCP for AI-powered conversations or CLI for programmatic control - no Excel programming knowledge required.

<div class="callout">
💡 <strong>Interactive Development</strong><br>
Unlike file-based tools, ExcelMcp lets you see results instantly in Excel - create → test → refine → iterate in real-time.
</div>

<div class="callout">
💻 <strong>For Developers</strong><br>
Think of Excel as an AI-powered REPL - write code (Power Query M, DAX, VBA), execute instantly, inspect results visually in the live workbook.
</div>

## Key Features

<div class="features-grid">
<div class="feature-card">
<h3>165 Operations</h3>
<p>11 specialized tools covering Power Query, DAX, VBA, PivotTables, ranges, formatting, and more</p>
</div>

<div class="feature-card">
<h3>100% Safe</h3>
<p>Uses Excel's native COM API - zero corruption risk, no file parsing</p>
</div>

<div class="feature-card">
<h3>Interactive Development</h3>
<p>See changes in real-time - create, test, refine, and iterate instantly</p>
</div>

<div class="feature-card">
<h3>Natural Language Control</h3>
<p>Describe what you want, AI does the rest</p>
</div>

<div class="feature-card">
<h3>Comprehensive Automation</h3>
<p>Power Query, Power Pivot, VBA, Tables, ranges, formatting</p>
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

**Requirements:** Windows OS, Microsoft Excel 2016+, .NET 8.0 Runtime

📖 **[Complete Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md)**

## Documentation

- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- [Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md)
- [Contributing Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/CONTRIBUTING.md)
- [API Reference](https://github.com/sbroenne/mcp-server-excel/tree/main/docs)

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
<li><strong>Comprehensive:</strong> 165 operations across 11 tools</li>
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
Same 165 operations available via CLI and MCP. Consistent behavior across interfaces. Full .NET library available for custom integrations.
</div>

## Related Projects

- [Model Context Protocol](https://modelcontextprotocol.io/) - The protocol specification
- [MCP Servers](https://github.com/modelcontextprotocol/servers) - Official MCP server examples
- [Claude Desktop](https://claude.ai/download) - AI assistant with MCP support
- [GitHub Copilot](https://github.com/features/copilot) - AI coding assistant

## Keywords

<div class="keywords">
Excel automation, MCP server, Model Context Protocol, GitHub Copilot, Claude AI, Power Query, M language, DAX, Power Pivot, VBA macros, Excel Tables, PivotTables, data analysis, spreadsheet automation, RPA, robotic process automation, COM automation, Windows automation, AI Excel assistant, natural language Excel, Excel API, Excel COM interop, Excel CLI, command line Excel, Excel scripting, Excel batch processing, CI/CD Excel, DevOps Excel, Excel PowerShell, Excel automation tool, .NET Excel library, Excel NuGet package
</div>

</div>

<footer>
<div class="container">
<p><strong>Excel MCP Server</strong> - MIT License</p>
<p><a href="https://github.com/sbroenne/mcp-server-excel">GitHub Repository</a> | <a href="https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md">Installation Guide</a> | <a href="https://github.com/sbroenne/mcp-server-excel/blob/main/docs/CONTRIBUTING.md">Contributing</a></p>
</div>
</footer>
