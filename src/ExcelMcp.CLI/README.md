# ExcelMcp.CLI - Command-Line Interface for Excel Automation

[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.CLI.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)
[![Downloads](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.CLI.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**A professional command-line tool for Excel development workflows, Power Query management, VBA automation, and data operations.**

Control Microsoft Excel from your terminal - manage worksheets, Power Query M code, DAX measures, PivotTables, Excel Tables, VBA macros, and more. Perfect for CI/CD pipelines, automated testing, and reproducible Excel workflows.

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

### Update to Latest Version

```bash
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

### Uninstall

```bash
dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI
```

---

## ✨ Key Features

### 🔧 Excel Development Automation
- **Power Query Management** - Export, import, update, and version control M code
- **VBA Development** - Manage VBA modules, run macros, automated testing
- **Data Model & DAX** - Create measures, manage relationships, Power Pivot operations
- **PivotTable Automation** - Create, configure, and manage PivotTables programmatically

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

ExcelMcp.CLI provides **147 commands** across 12 categories:

| Category | Commands | Examples |
|----------|----------|----------|
| **File Operations** | 1 | `create-empty` |
| **Worksheets** | 13 | `sheet-list`, `sheet-create`, `sheet-rename`, `sheet-set-tab-color` |
| **Power Query** | 9 | `pq-list`, `pq-create`, `pq-export`, `pq-refresh`, `pq-update-mcode` |
| **Ranges** | 44 | `range-get-values`, `range-set-values`, `range-copy`, `range-find`, `range-merge-cells`, `range-add-hyperlink` |
| **Excel Tables** | 23 | `table-create`, `table-filter`, `table-sort`, `table-add-column`, `table-get-column-format` |
| **PivotTables** | 12 | `pivot-create-from-range`, `pivot-add-row-field`, `pivot-refresh`, `pivot-delete` |
| **QueryTables** | 8 | `querytable-list`, `querytable-get`, `querytable-refresh`, `querytable-create-from-connection` |
| **Data Model** | 15 | `dm-create-measure`, `dm-create-relationship`, `dm-refresh` |
| **Connections** | 11 | `conn-list`, `conn-import`, `conn-refresh`, `conn-test` |
| **Named Ranges** | 6 | `namedrange-create`, `namedrange-get`, `namedrange-set` |
| **VBA** | 5 | `vba-list`, `vba-import`, `vba-run`, `vba-export` |

> **Note:** Recent expansion added 38 new commands (33 Range operations, 3 QueryTable methods, 2 Table methods). 7 PivotTable operations are planned for future releases.

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
