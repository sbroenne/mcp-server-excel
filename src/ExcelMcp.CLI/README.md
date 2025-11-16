# ExcelMcp.CLI - Command-Line Interface for Excel Automation

[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.CLI.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)
[![Downloads](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.CLI.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**A professional command-line tool for Excel development workflows, Power Query management, VBA automation, and data operations.**

Control Microsoft Excel from your terminal - manage worksheets, Power Query M code, DAX measures, PivotTables, Excel Tables, VBA macros, and more. Perfect for CI/CD pipelines, automated testing, and reproducible Excel workflows.

➡️ **[Learn more and see examples](https://sbroenne.github.io/mcp-server-excel/)**

---

## 🚀 Quick Start

### Installation (.NET Global Tool - Recommended)

```bash
# Install globally (requires .NET 8 SDK)
dotnet tool install --global Sbroenne.ExcelMcp.CLI

# Verify installation
excelcli --version

# Get help
excelcli --help
```

> 🔁 **Session Workflow:** Always start with `excelcli session open <file>` (captures the session id), pass `--session <id>` to other commands, then `excelcli session save <id>` (optional) and `excelcli session close <id>` when finished. The CLI reuses the same Excel instance through that lifecycle.

### Update to Latest Version

```bash
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

### Uninstall

```bash
dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI
```

## 🆘 Built-in Help

- `excelcli --help` – lists every command category plus the new descriptions from `Program.cs`
- `excelcli <command> --help` – shows verb-specific arguments (for example `excelcli sheet --help`)
- `excelcli session --help` – displays nested verbs such as `open`, `save`, `close`, and `list`

Descriptions are kept in sync with the CLI source so the help output always reflects the latest capabilities.

---

## ✨ Key Features

### 🔧 Excel Development Automation
- **Power Query Management** - Export, import, update, and version control M code
- **VBA Development** - Manage VBA modules, run macros, automated testing
- **Data Model & DAX** - Create measures, manage relationships, Power Pivot operations
- **PivotTable Automation** - Create, configure, and manage PivotTables programmatically
- **Conditional Formatting** - Add rules (cell value, expression-based), clear formatting

### 📊 Data Operations
- **Worksheet Management** - Create, rename, copy, delete sheets with tab colors and visibility
- **Range Operations** - Read/write values, formulas, formatting, validation
- **Excel Tables** - Lifecycle management, filtering, sorting, structured references
- **Connection Management** - OLEDB, ODBC, Text, Web connections with testing

### 🛡️ Production Ready
- **Zero Corruption Risk** - Uses Excel's native COM API (not file manipulation)
- **Error Handling** - Comprehensive validation and helpful error messages
- **CI/CD Integration** - Perfect for automated workflows and testing
- **Windows Native** - Optimized for Windows Excel automation

---

## 📋 Command Categories

ExcelMcp.CLI provides **168 operations** across 12 categories:

| Category | Operations | Examples |
|----------|-----------|----------|
| **File & Session** | 5 | `create-empty`, `session open`, `session save`, `session close`, `session list` |
| **Worksheets** | 13 | `sheet-list`, `sheet-create`, `sheet-rename`, `sheet-set-tab-color` |
| **Power Query** | 12 | `pq-list`, `pq-create`, `pq-export`, `pq-refresh`, `pq-update-mcode` |
| **Ranges** | 43 | `range-get-values`, `range-set-values`, `range-copy`, `range-find`, `range-merge-cells`, `range-add-hyperlink` |
| **Conditional Formatting** | 2 | `cf-add-rule`, `cf-clear-rules` |
| **Excel Tables** | 23 | `table-create`, `table-filter`, `table-sort`, `table-add-column`, `table-get-column-format` |
| **PivotTables** | 19 | `pivot-create-from-range`, `pivot-add-row-field`, `pivot-refresh`, `pivot-delete` |
| **QueryTables** | 8 | `querytable-list`, `querytable-get`, `querytable-refresh`, `querytable-create-from-connection` |
| **Data Model** | 15 | `dm-create-measure`, `dm-create-relationship`, `dm-refresh` |
| **Connections** | 12 | `conn-list`, `conn-import`, `conn-refresh`, `conn-test` |
| **Named Ranges** | 7 | `namedrange-create`, `namedrange-get`, `namedrange-set`, `namedrange-create-bulk` |
| **VBA** | 7 | `vba-list`, `vba-import`, `vba-run`, `vba-export`

---

## SESSION LIFECYCLE (Open/Save/Close)

The CLI uses an explicit session-based workflow where you open a file, perform operations, and save or close:

```bash
# 1. Open a session
excelcli open data.xlsx
# Output: Session ID: 550e8400-e29b-41d4-a716-446655440000

# 2. List active sessions anytime
excelcli list

# 3. Use the session ID with any commands (optional - can operate without session)
excelcli sheet-create data.xlsx NewSheet --session-id 550e8400-e29b-41d4-a716-446655440000
excelcli pq-list data.xlsx --session-id 550e8400-e29b-41d4-a716-446655440000

# 4. Save changes and keep session open
excelcli save 550e8400-e29b-41d4-a716-446655440000

# 5. Close session and discard changes
excelcli close 550e8400-e29b-41d4-a716-446655440000
```

### Session Lifecycle Benefits

- **Explicit control** - Know exactly when changes are persisted
- **Batch efficiency** - Keep single Excel instance open for multiple operations (75-90% faster)
- **Flexibility** - Save strategically or discard changes entirely
- **Clean resource management** - Automatic Excel cleanup when session closes

---

## 💡 Common Use Cases

### Power Query Development

```bash
# Export all queries for version control
excelcli pq-list data.xlsx
excelcli pq-export data.xlsx "Sales Data" sales-query.pq
excelcli pq-export data.xlsx "Customer Data" customer-query.pq

# Import updated query from file
excelcli pq-import data.xlsx "Sales Data" sales-query-optimized.pq

# Refresh all queries
excelcli pq-refresh data.xlsx "Sales Data"
```

### VBA Module Management

```bash
# List all VBA modules
excelcli vba-list macros.xlsm

# Export module for version control
excelcli vba-export macros.xlsm "DataProcessor" processor.vba

# Import updated module
excelcli vba-import macros.xlsm "DataProcessor" processor-v2.vba

# Run a macro with parameters
excelcli vba-run macros.xlsm "ProcessData" "Sheet1" "A1:D100"
```

### Data Model & DAX

```bash
# Create a DAX measure
excelcli dm-create-measure sales.xlsx Sales "TotalRevenue" "SUM(Sales[Amount])" Currency

# Create relationship between tables
excelcli dm-create-relationship sales.xlsx Sales CustomerID Customers ID

# List all measures
excelcli dm-list-measures sales.xlsx

# Refresh Data Model
excelcli dm-refresh sales.xlsx
```

### Excel Table Operations

```bash
# Create table from range
excelcli table-create data.xlsx Sheet1 SalesTable A1:E100

# Apply filters
excelcli table-apply-filter data.xlsx SalesTable Amount ">1000"
excelcli table-apply-filter-values data.xlsx SalesTable Region "North,South,East"

# Sort by column
excelcli table-sort data.xlsx SalesTable Amount desc

# Add calculated column
excelcli table-add-column data.xlsx SalesTable "Total" 5
```

### PivotTable Automation

```bash
# Create PivotTable from range
excelcli pivot-create-from-range sales.xlsx Data A1:D100 Analysis A1 SalesPivot

# Configure fields
excelcli pivot-add-row-field sales.xlsx SalesPivot Region
excelcli pivot-add-column-field sales.xlsx SalesPivot Year
excelcli pivot-add-value-field sales.xlsx SalesPivot Amount Sum "Total Sales"
excelcli pivot-add-filter-field sales.xlsx SalesPivot Category

# Manage PivotTable lifecycle
excelcli pivot-get sales.xlsx SalesPivot
excelcli pivot-list-fields sales.xlsx SalesPivot
excelcli pivot-remove-field sales.xlsx SalesPivot Year
excelcli pivot-delete sales.xlsx SalesPivot

# Create from Data Model (for large datasets)
excelcli pivot-create-from-datamodel sales.xlsx ConsumptionMilestones Analysis A1 MilestonesPivot

# Refresh PivotTable
excelcli pivot-refresh sales.xlsx SalesPivot
```

### Session Mode for RPA Workflows

```bash
# Example: Automated report generation with session lifecycle

# 1. Open session
SESSION_ID=$(excelcli open report.xlsx | grep "Session ID:" | cut -d' ' -f3)

# 2. Perform operations (all use same Excel instance)
excelcli sheet-create report.xlsx "Sales" --session-id $SESSION_ID
excelcli sheet-create report.xlsx "Customers" --session-id $SESSION_ID
excelcli sheet-create report.xlsx "Summary" --session-id $SESSION_ID

# 3. Import data
excelcli range-set-values report.xlsx Sales A1 "sales.csv" --session-id $SESSION_ID
excelcli range-set-values report.xlsx Customers A1 "customers.csv" --session-id $SESSION_ID

# 4. Add Power Query for transformations
excelcli pq-create report.xlsx "CleanSales" "clean-sales.pq" --session-id $SESSION_ID

# 5. Create PivotTable
excelcli pivot-create-from-range report.xlsx Sales A1:E1000 Summary A1 SalesPivot --session-id $SESSION_ID

# 6. Save changes
excelcli save $SESSION_ID

# 7. Close session
excelcli close $SESSION_ID
```

### QueryTable Operations

```bash
# List all QueryTables in workbook
excelcli querytable-list data.xlsx

# Get QueryTable details
excelcli querytable-get data.xlsx "WebData"

# Refresh QueryTables
excelcli querytable-refresh data.xlsx "WebData"
excelcli querytable-refresh-all data.xlsx

# Delete QueryTable
excelcli querytable-delete data.xlsx "WebData"
```

### Worksheet Management

```bash
# List all sheets
excelcli sheet-list workbook.xlsx

# Create and configure sheets
excelcli sheet-create workbook.xlsx "Q1 Data"
excelcli sheet-set-tab-color workbook.xlsx "Q1 Data" 0 255 0  # Green
excelcli sheet-hide workbook.xlsx "Calculations"

# Rename sheets
excelcli sheet-rename workbook.xlsx "Sheet1" "Sales Summary"
```

### Range Operations

```bash
# Read range data
excelcli range-get-values data.xlsx Sheet1 A1:D10

# Write CSV data to range
excelcli range-set-values data.xlsx Sheet1 A1:C10 data.csv

# Apply formatting
excelcli range-format data.xlsx Sheet1 A1:E1 --bold --font-size 12 --h-align Center
excelcli range-format data.xlsx Sheet1 D2:D100 --fill-color "#FFFF00"

# Set number formats
excelcli range-set-number-format data.xlsx Sheet1 D2:D100 "$#,##0.00"  # Currency
excelcli range-set-number-format data.xlsx Sheet1 E2:E100 "0.00%"      # Percentage

# Add data validation
excelcli range-validate data.xlsx Sheet1 F2:F100 List "Active,Inactive,Pending"
```

### Conditional Formatting

```bash
# Add conditional formatting rule (highlight cells > 100)
excelcli cf-add-rule data.xlsx Sheet1 A1:A10 cell-value greater 100 "" "#FFFF00" solid

# Add expression-based rule
excelcli cf-add-rule data.xlsx Sheet1 B1:B10 expression "" "=B1>AVERAGE($B$1:$B$10)" "" "#90EE90" solid

# Clear conditional formatting
excelcli cf-clear-rules data.xlsx Sheet1 A1:A10
```

---

## ⚙️ System Requirements

| Requirement | Details | Why Required |
|-------------|---------|--------------|
| **Windows OS** | Windows 10/11 or Server 2016+ | COM interop is Windows-specific |
| **Microsoft Excel** | Excel 2016 or later | CLI controls actual Excel application |
| **.NET 8 Runtime** | [Download](https://dotnet.microsoft.com/download/dotnet/8.0) | Required to run .NET global tools |

> **Note:** ExcelMcp.CLI controls the actual Excel application via COM interop, not just file formats. This provides access to Power Query, VBA runtime, formula engine, and all Excel features, but requires Excel to be installed.

---

## 🔒 VBA Operations Setup (One-Time)

VBA commands require **"Trust access to the VBA project object model"** to be enabled:

1. Open Excel
2. Go to **File → Options → Trust Center**
3. Click **"Trust Center Settings"**
4. Select **"Macro Settings"**
5. Check **"✓ Trust access to the VBA project object model"**
6. Click **OK** twice

This is a security setting that must be manually enabled. ExcelMcp.CLI never modifies security settings automatically.

For macro-enabled workbooks, use `.xlsm` extension:

```bash
excelcli create-empty macros.xlsm
excelcli vba-import macros.xlsm "Module1" code.vba
```

---

## 📖 Complete Documentation

- **[NuGet Package](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)** - .NET Global Tool installation
- **[GitHub Repository](https://github.com/sbroenne/mcp-server-excel)** - Source code and issues
- **[Release Notes](https://github.com/sbroenne/mcp-server-excel/releases)** - Latest updates

---

## 🚧 Troubleshooting

### Command Not Found After Installation

```bash
# Verify .NET tools path is in your PATH environment variable
dotnet tool list --global

# If excelcli is listed but not found, add .NET tools to PATH:
# The default location is: %USERPROFILE%\.dotnet\tools
```

### Excel Not Found

```bash
# Error: "Microsoft Excel is not installed"
# Solution: Install Microsoft Excel (any version 2016+)
```

### VBA Access Denied

```bash
# Error: "Programmatic access to Visual Basic Project is not trusted"
# Solution: Enable VBA trust (see VBA Operations Setup above)
```

### Permission Issues

```bash
# Run PowerShell/CMD as Administrator if you encounter permission errors
# Or install to user directory: dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

---

## 🛠️ Advanced Usage

### Scripting & Automation

```bash
# PowerShell script example
$files = Get-ChildItem *.xlsx
foreach ($file in $files) {
    excelcli pq-refresh $file.Name "Sales Data"
    excelcli dm-refresh $file.Name
}
```

### CI/CD Integration

```yaml
# GitHub Actions example
- name: Install ExcelMcp.CLI
  run: dotnet tool install --global Sbroenne.ExcelMcp.CLI

- name: Process Excel Files
  run: |
    excelcli pq-import data.xlsx "Query1" queries/query1.pq
    excelcli pq-refresh data.xlsx "Query1"
```

### Batch Processing

```bash
# Process multiple files
for %f in (*.xlsx) do excelcli sheet-read "%f" "Sheet1" >> output.csv
```

## ✅ Tested Scenarios

The CLI ships with real Excel-backed integration tests that exercise the session lifecycle plus worksheet creation/listing flows through the same commands you run locally. Execute them with:

```bash
dotnet test tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj --filter "Layer=CLI"
```

These tests open actual workbooks, issue `session open/list/close`, and call `excelcli sheet` actions to ensure the command pipeline stays healthy.

---

## 🤝 Related Tools

- **[ExcelMcp.McpServer](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)** - MCP server for AI assistant integration
- **[Excel MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)** - One-click Excel automation in VS Code
- **[ExcelMcp.Core](https://www.nuget.org/packages/Sbroenne.ExcelMcp.Core)** - Core library for custom automation tools

---

## 📄 License

MIT License - see [LICENSE](../../LICENSE) for details.

---

## 🙋 Support

- **Issues**: [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **Discussions**: [GitHub Discussions](https://github.com/sbroenne/mcp-server-excel/discussions)
- **Documentation**: [Complete Docs](../../docs/)

---

**Built with ❤️ for Excel developers and automation engineers**
