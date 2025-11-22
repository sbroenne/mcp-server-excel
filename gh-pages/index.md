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

<div class="badges-section">
  <div class="container">
    <div class="hero-badges">
      <a href="https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer"><img src="https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.McpServer.svg?label=NuGet%20MCP%20Installs" alt="NuGet MCP Server Installs"></a>
      <a href="https://github.com/sbroenne/mcp-server-excel/releases"><img src="https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=GitHub%20Downloads" alt="GitHub Downloads"></a>
    </div>
  </div>
</div>

<div class="container content-section" markdown="1">
## 🤔 What is This?

Stop manually clicking through Excel menus for repetitive tasks. Instead, describe what you want in plain English and let the coding agent and the MCP Server handle the task:

- *"Load this sales data CSV file into Excel. Only keep the columns that I need to compare sales numbers month over month. Load them to a data model and create the necessary DAX measures. Add pivot tables.*
- *"Create a PivotTable from SalesData table showing top 10 products by region with sum and average"*
- *"Apply conditional formatting to highlight values above $10,000 in red and below $5,000 in yellow"*
- *"Convert this data range to an Excel Table with style TableStyleMedium2, add auto-filters, and create a totals row"*
- *"Add data validation dropdowns to the Status column with options: Active, Pending, Completed"*
- *"Merge the header cells, center-align them, and auto-fit all column widths to content"*
- "Extract all PowerQueries, DAX measures and VBA code so I can use version control in GIT".

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
- ✅ **Comprehensive automation** - Currently supports 154 operations across 11 specialized tools covering Power Query, Data Model/DAX, VBA, PivotTables, Excel Tables, ranges, conditional formatting, and more

**💻 For Developers:** Think of Excel as an AI-powered REPL - write code (Power Query M, DAX, VBA), execute instantly, inspect results visually in the live workbook. No more blind editing of .xlsx files.

## Key Features

<div class="features-grid">
<div class="feature-card">
<h3>154 Operations</h3>
<p>11 specialized tools (154 operations) covering Power Query, DAX, VBA, PivotTables, ranges, conditional formatting, and more</p>
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
<p>154 operations covering Power Query, DAX, VBA, Tables, and more</p>
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
<li><strong>Comprehensive:</strong> 154 operations across 11 tools</li>
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
Same 154 operations available via CLI and MCP. Consistent behavior across interfaces. Full .NET library available for custom integrations.
</div>

## Complete Tool & Action Reference

**11 specialized tools with 154 operations:**

<div class="tools-reference">

### 📊 Power Query & M Code (9 actions)

| Action | Description |
|--------|-------------|
| `list` | List all Power Query queries in workbook |
| `view` | View M code for a specific query |
| `create` | Create new Power Query from M code |
| `update` | Update existing query's M code |
| `delete` | Delete a Power Query |
| `refresh` | Refresh a specific query to reload data |
| `refresh-all` | Refresh all queries in workbook |
| `get-load-config` | Get load destination settings for a query |
| `load-to` | Load query data to worksheet, data model, or both |

### 🔢 Power Pivot / Data Model (14 actions)

| Action | Description |
|--------|-------------|
| `list-tables` | List all tables in Data Model |
| `read-table` | Read data from a Data Model table |
| `list-columns` | List all columns in a table |
| `list-measures` | List all DAX measures in a table |
| `read` | Read details of a specific measure |
| `create-measure` | Create new DAX measure |
| `update-measure` | Update existing DAX measure formula or format |
| `delete-measure` | Delete a DAX measure |
| `list-relationships` | List all table relationships |
| `create-relationship` | Create relationship between tables |
| `update-relationship` | Update relationship properties |
| `delete-relationship` | Delete a relationship |
| `read-info` | Get Data Model metadata and statistics |
| `refresh` | Refresh entire Data Model |

### 📋 Excel Tables (23 actions)

| Action | Description |
|--------|-------------|
| `list` | List all Excel Tables in workbook |
| `read` | Read table data and properties |
| `create` | Create new Excel Table from range |
| `rename` | Rename an Excel Table |
| `delete` | Delete an Excel Table |
| `resize` | Resize table range |
| `set-style` | Apply table style (TableStyleMedium2, etc.) |
| `toggle-totals` | Show/hide totals row |
| `set-column-total` | Set aggregation function for column total |
| `append` | Append rows to table |
| `add-to-datamodel` | Add table to Power Pivot Data Model |
| `apply-filter` | Apply filter criteria to column |
| `apply-filter-values` | Filter by specific values |
| `clear-filters` | Clear all filters |
| `get-filters` | Get current filter settings |
| `add-column` | Add new column to table |
| `remove-column` | Remove column from table |
| `rename-column` | Rename table column |
| `get-structured-reference` | Get structured reference formula |
| `sort` | Sort table by column |
| `sort-multi` | Sort by multiple columns |
| `get-column-number-format` | Get number format for column |
| `set-column-number-format` | Set number format for column |

### 📈 PivotTables (25 actions)

| Action | Description |
|--------|-------------|
| `list` | List all PivotTables in workbook |
| `read` | Read PivotTable configuration |
| `create-from-range` | Create PivotTable from range |
| `create-from-table` | Create PivotTable from Excel Table |
| `create-from-datamodel` | Create PivotTable from Data Model |
| `delete` | Delete a PivotTable |
| `refresh` | Refresh PivotTable data |
| `list-fields` | List all available fields |
| `add-row-field` | Add field to Rows area |
| `add-column-field` | Add field to Columns area |
| `add-value-field` | Add field to Values area with aggregation |
| `add-filter-field` | Add field to Filters area |
| `remove-field` | Remove field from PivotTable |
| `set-field-function` | Set aggregation function (Sum, Count, Average, etc.) |
| `set-field-name` | Set custom field display name |
| `set-field-format` | Set number format for value field |
| `set-field-filter` | Apply filter to a field |
| `sort-field` | Sort field values |
| `get-data` | Extract PivotTable data as values |
| `set-grand-totals` | Configure grand totals display |
| `set-column-grand-totals` | Show/hide column grand totals |
| `set-row-grand-totals` | Show/hide row grand totals |
| `get-grand-totals` | Get grand totals configuration |
| `set-subtotals` | Configure subtotals for fields |
| `get-subtotals` | Get subtotals configuration |

### 📝 Ranges & Data (42 actions)

| Action | Description |
|--------|-------------|
| `get-values` | Read cell values from range |
| `set-values` | Write values to range |
| `get-formulas` | Read formulas from range |
| `set-formulas` | Write formulas to range |
| `get-number-formats` | Get number formats for range |
| `set-number-format` | Set number format (currency, percentage, date, etc.) |
| `set-number-formats` | Set multiple number formats at once |
| `clear-all` | Clear values, formulas, and formatting |
| `clear-contents` | Clear values and formulas only |
| `clear-formats` | Clear formatting only |
| `copy` | Copy range (all attributes) |
| `copy-values` | Copy values only |
| `copy-formulas` | Copy formulas only |
| `insert-cells` | Insert cells with shift direction |
| `delete-cells` | Delete cells with shift direction |
| `insert-rows` | Insert rows |
| `delete-rows` | Delete rows |
| `insert-columns` | Insert columns |
| `delete-columns` | Delete columns |
| `find` | Find cells matching criteria |
| `replace` | Find and replace in range |
| `sort` | Sort range by columns |
| `get-used-range` | Get actual used range in worksheet |
| `get-current-region` | Get contiguous range around cell |
| `get-info` | Get range metadata (address, size, etc.) |
| `add-hyperlink` | Add hyperlink to cell |
| `remove-hyperlink` | Remove hyperlink from cell |
| `list-hyperlinks` | List all hyperlinks in range |
| `get-hyperlink` | Get hyperlink details |
| `get-style` | Get cell style name |
| `set-style` | Apply built-in Excel style |
| `format-range` | Apply visual formatting (font, fill, border, alignment) |
| `validate-range` | Add data validation rules (dropdowns, number/date/text rules) |
| `get-validation` | Get validation settings |
| `remove-validation` | Remove data validation |
| `autofit-columns` | Auto-fit column widths |
| `autofit-rows` | Auto-fit row heights |
| `merge-cells` | Merge cells in range |
| `unmerge-cells` | Unmerge cells |
| `get-merge-info` | Get merge status and areas |
| `set-cell-lock` | Lock/unlock cells for protection |
| `get-cell-lock` | Get cell lock status |

### 🎨 Conditional Formatting (2 actions)

| Action | Description |
|--------|-------------|
| `add-rule` | Add conditional formatting rule (cell value or expression-based) |
| `clear-rules` | Clear conditional formatting from range |

### 🖥️ VBA Macros (6 actions)

| Action | Description |
|--------|-------------|
| `list` | List all VBA modules in workbook |
| `view` | View VBA module code |
| `import` | Import VBA code from file |
| `delete` | Delete VBA module |
| `run` | Execute VBA macro |
| `update` | Update existing VBA module code |

### 🔌 Data Connections (9 actions)

| Action | Description |
|--------|-------------|
| `list` | List all data connections |
| `view` | View connection details and settings |
| `create` | Create new data connection (OLEDB, ODBC, Text, Web) |
| `test` | Test connection validity |
| `refresh` | Refresh connection to reload data |
| `delete` | Delete a connection |
| `load-to` | Load connection data to worksheet |
| `get-properties` | Get connection properties |
| `set-properties` | Update connection properties |

### 📄 Worksheets (16 actions)

| Action | Description |
|--------|-------------|
| `list` | List all worksheets in workbook |
| `create` | Create new worksheet |
| `rename` | Rename worksheet |
| `copy` | Copy worksheet within same workbook |
| `delete` | Delete worksheet |
| `move` | Move worksheet to different position |
| `copy-to-workbook` | Copy worksheet to different workbook |
| `move-to-workbook` | Move worksheet to different workbook |
| `set-tab-color` | Set worksheet tab color (RGB) |
| `get-tab-color` | Get current tab color |
| `clear-tab-color` | Remove tab color |
| `hide` | Hide worksheet from UI (visible in VBA) |
| `very-hide` | Hide from UI and VBA |
| `show` | Make worksheet visible |
| `get-visibility` | Get current visibility state |
| `set-visibility` | Set visibility state |

### 🏷️ Named Ranges (7 actions)

| Action | Description |
|--------|-------------|
| `list` | List all named ranges |
| `read` | Read named range value |
| `write` | Write value to named range |
| `create` | Create new named range |
| `create-bulk` | Create multiple named ranges at once |
| `update` | Update named range reference |
| `delete` | Delete named range |

### 📁 File & Batch Operations (6 actions)

| Action | Description |
|--------|-------------|
| `open` | Open workbook session |
| `save` | Save changes to workbook |
| `close` | Close workbook session |
| `create-empty` | Create new empty workbook (.xlsx or .xlsm) |
| `test` | Test if workbook can be opened |
| `begin-batch` | Start batch session for multiple operations (75-90% faster) |

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
