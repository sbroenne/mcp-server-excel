# ExcelMcp.CLI - Command-Line Interface for Excel Automation

[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.CLI.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)
[![Downloads](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.CLI.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.CLI)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Optional command-line interface for Excel automation without AI assistance.**

For advanced users who prefer direct scripting control, the CLI provides 13 command categories with 194 operations matching the MCP Server (22 tools with 194 operations). Perfect for RPA workflows, CI/CD pipelines, batch processing, and automated testing.

**Note:** Most users should use the [MCP Server with AI assistants](../ExcelMcp.McpServer/README.md) for natural language automation.

➡️ **[Learn more and see examples](https://sbroenne.github.io/mcp-server-excel/)**

---

## 🚀 Quick Start

### Installation (.NET Global Tool - Recommended)

```bash
# Install globally (requires .NET 10 SDK)
dotnet tool install --global Sbroenne.ExcelMcp.CLI

# Verify installation
excelcli --version

# Get help
excelcli --help
```

> 🔁 **Session Workflow:** Always start with `excelcli session open <file>` (captures the session id), pass `--session-id <id>` to other commands, then `excelcli session save <id>` (optional) and `excelcli session close <id>` when finished. The CLI reuses the same Excel instance through that lifecycle.

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

ExcelMcp.CLI provides **194 operations** across 13 categories:

📚 **[Complete Feature Reference →](../../FEATURES.md)** - Full MCP Server documentation (194 operations)

**Quick Reference:**

| Category | Operations | Examples |
|----------|-----------|----------|
| **File & Session** | 5 | `create-empty`, `session open`, `session save`, `session close`, `session list` |
| **Worksheets** | 16 | `sheet list`, `sheet create`, `sheet rename`, `sheet copy`, `sheet move`, `sheet copy-to-file`, `sheet move-to-file`, `sheet set-tab-color`, `sheet get-visibility` |
| **Power Query** | 10 | `powerquery list`, `powerquery create`, `powerquery refresh`, `powerquery update`, `powerquery rename` |
| **Ranges** | 42 | `range get-values`, `range set-values`, `range copy`, `range find`, `range merge-cells`, `range add-hyperlink` |
| **Conditional Formatting** | 2 | `conditionalformat add-rule`, `conditionalformat clear-rules` |
| **Excel Tables** | 27 | `table create`, `table apply-filter`, `table get-data`, `table sort`, `table add-column`, `table create-from-dax`, `table update-dax` |
| **Charts** | 14 | `chart create-from-range`, `chart add-series`, `chart set-chart-type`, `chart show-legend` |
| **PivotTables** | 30 | `pivottable create-from-range`, `pivottable add-row-field`, `pivottable refresh`, `pivottable list-calculated-fields`, `pivottable create-calculated-member` |
| **Data Model** | 18 | `datamodel create-measure`, `datamodel create-relationship`, `datamodel refresh`, `datamodel evaluate`, `datamodel delete-table`, `datamodel read-relationship` |
| **Connections** | 9 | `connection list`, `connection refresh`, `connection test` |
| **Named Ranges** | 6 | `namedrange create`, `namedrange read`, `namedrange write`, `namedrange update` |
| **VBA** | 7 | `vba list`, `vba import`, `vba run`, `vba update` |

**Note:** CLI uses session commands for multi-operation workflows.

---

## SESSION LIFECYCLE (Open/Save/Close)

The CLI uses an explicit session-based workflow where you open a file, perform operations, and save or close:

```bash
# 1. Open a session
excelcli session open data.xlsx
# Output: Session ID: 550e8400-e29b-41d4-a716-446655440000

# 2. List active sessions anytime
excelcli session list

# 3. Use the session ID with any commands
excelcli sheet create --session-id 550e8400-e29b-41d4-a716-446655440000 --sheet "NewSheet"
excelcli powerquery list --session-id 550e8400-e29b-41d4-a716-446655440000

# 4. Save changes and keep session open
excelcli session save 550e8400-e29b-41d4-a716-446655440000

# 5. Close session and discard changes
excelcli session close 550e8400-e29b-41d4-a716-446655440000
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
# List all queries
excelcli powerquery list --session-id <SESSION>

# View a query
excelcli powerquery view --session-id <SESSION> --query "Sales Data"

# Create a query from M code file
excelcli powerquery create --session-id <SESSION> --query "Sales Data" --m-file sales-query.pq

# Update existing query
excelcli powerquery update --session-id <SESSION> --query "Sales Data" --m-file sales-query-optimized.pq

# Rename a query
excelcli powerquery rename --session-id <SESSION> --query "Sales Data" --new-name "Sales Data v2"

# Refresh a query
excelcli powerquery refresh --session-id <SESSION> --query "Sales Data"

# Refresh all queries
excelcli powerquery refresh-all --session-id <SESSION>
```

### VBA Module Management

```bash
# List all VBA modules
excelcli vba list --session-id <SESSION>

# View a module
excelcli vba view --session-id <SESSION> --module "DataProcessor"

# Export module for version control
excelcli vba export --session-id <SESSION> --module "DataProcessor" --output processor.vba

# Import updated module
excelcli vba import --session-id <SESSION> --module "DataProcessor" --input processor-v2.vba

# Update existing module
excelcli vba update --session-id <SESSION> --module "DataProcessor" --input processor-updated.vba

# Run a macro
excelcli vba run --session-id <SESSION> --procedure "Module1.ProcessData"
```

### Data Model & DAX

```bash
# List all tables
excelcli datamodel list-tables --session-id <SESSION>

# List all measures in a table
excelcli datamodel list-measures --session-id <SESSION> --table Sales

# Create a DAX measure
excelcli datamodel create-measure --session-id <SESSION> --table Sales --name "TotalRevenue" --formula "SUM(Sales[Amount])" --format Currency

# Update a measure
excelcli datamodel update-measure --session-id <SESSION> --table Sales --name "TotalRevenue" --formula "SUM(Sales[Amount])" --format Currency

# Rename a Data Model table (Power Query-backed tables only)
excelcli datamodel rename-table --session-id <SESSION> --table "Sales" --new-name "SalesData"

# Create relationship between tables
excelcli datamodel create-relationship --session-id <SESSION> --from-table Sales --from-column CustomerID --to-table Customers --to-column ID

# Refresh Data Model
excelcli datamodel refresh --session-id <SESSION>

# Execute DAX EVALUATE query
excelcli datamodel evaluate --session-id <SESSION> --dax-query "EVALUATE SUMMARIZE(Sales, Sales[Region], 'Total', SUM(Sales[Amount]))"
```

### Excel Table Operations

```bash
# List all tables
excelcli table list --session-id <SESSION>

# Create table from range
excelcli table create --session-id <SESSION> --sheet Sheet1 --table-name SalesTable --range A1:E100

# Apply filter criteria
excelcli table apply-filter --session-id <SESSION> --table-name SalesTable --column Amount --criteria ">1000"

# Apply filter by values
excelcli table apply-filter-values --session-id <SESSION> --table-name SalesTable --column Region --values "North,South,East"

# Create a DAX-backed table from a query
excelcli table create-from-dax --session-id <SESSION> --sheet Results --table-name Summary --dax-query "EVALUATE SUMMARIZE(Sales, Sales[Region], 'Total', SUM(Sales[Amount]))"

# Update the DAX query for an existing table
excelcli table update-dax --session-id <SESSION> --table-name Summary --dax-query "EVALUATE TOPN(10, Sales, Sales[Amount], DESC)"

# Get the DAX query info for a table
excelcli table get-dax --session-id <SESSION> --table-name Summary

# Sort by column
excelcli table sort --session-id <SESSION> --table-name SalesTable --column Amount --descending

# Add column
excelcli table add-column --session-id <SESSION> --table-name SalesTable --column-name "Total" --position 5
```

### PivotTable Automation

```bash
# List all PivotTables
excelcli pivottable list --session-id <SESSION>

# Create PivotTable from range
excelcli pivottable create-from-range --session-id <SESSION> --source-sheet Data --source-range A1:D100 --dest-sheet Analysis --dest-cell A1 --name SalesPivot

# Create from Excel Table
excelcli pivottable create-from-table --session-id <SESSION> --table-name SalesData --dest-sheet Analysis --dest-cell A1 --name SalesPivot

# Create from Data Model
excelcli pivottable create-from-datamodel --session-id <SESSION> --table-name ConsumptionMilestones --dest-sheet Analysis --dest-cell A1 --name MilestonesPivot

# Configure fields
excelcli pivottable add-row-field --session-id <SESSION> --name SalesPivot --field Region
excelcli pivottable add-column-field --session-id <SESSION> --name SalesPivot --field Year
excelcli pivottable add-value-field --session-id <SESSION> --name SalesPivot --field Amount --function Sum --custom-name "Total Sales"
excelcli pivottable add-filter-field --session-id <SESSION> --name SalesPivot --field Category

# List fields
excelcli pivottable list-fields --session-id <SESSION> --name SalesPivot

# Remove field
excelcli pivottable remove-field --session-id <SESSION> --name SalesPivot --field Year

# Refresh PivotTable
excelcli pivottable refresh --session-id <SESSION> --name SalesPivot

# Delete PivotTable
excelcli pivottable delete --session-id <SESSION> --name SalesPivot
```

### Session Mode for RPA Workflows

```bash
# Example: Automated report generation with session lifecycle

# 1. Open session
SESSION_ID=$(excelcli session open report.xlsx | grep "Session ID:" | cut -d' ' -f3)

# 2. Perform operations (all use same Excel instance)
excelcli sheet create --session-id $SESSION_ID --sheet "Sales"
excelcli sheet create --session-id $SESSION_ID --sheet "Customers"
excelcli sheet create --session-id $SESSION_ID --sheet "Summary"

# 3. Import data
excelcli range set-values --session-id $SESSION_ID --sheet Sales --range A1 --values "[[...]]"

# 4. Add Power Query for transformations
excelcli powerquery create --session-id $SESSION_ID --query "CleanSales" --m-file "clean-sales.pq"

# 5. Create PivotTable
excelcli pivottable create-from-range --session-id $SESSION_ID --source-sheet Sales --source-range A1:E1000 --dest-sheet Summary --dest-cell A1 --name SalesPivot

# 6. Save changes
excelcli session save $SESSION_ID

# 7. Close session
excelcli session close $SESSION_ID
```

### Worksheet Management

```bash
# List all sheets
excelcli sheet list --session-id <SESSION>

# Create sheet
excelcli sheet create --session-id <SESSION> --sheet "Q1 Data"

# Rename sheet
excelcli sheet rename --session-id <SESSION> --sheet "Sheet1" --new-name "Sales Summary"

# Set tab color (RGB)
excelcli sheet set-tab-color --session-id <SESSION> --sheet "Q1 Data" --red 0 --green 255 --blue 0

# Hide sheet
excelcli sheet hide --session-id <SESSION> --sheet "Calculations"

# Show sheet
excelcli sheet show --session-id <SESSION> --sheet "Calculations"

# Copy sheet
excelcli sheet copy --session-id <SESSION> --sheet "Template" --new-name "Q1 Report"
```

### Range Operations

```bash
# Read range values
excelcli range get-values --session-id <SESSION> --sheet Sheet1 --range A1:D10

# Set range values (JSON array)
excelcli range set-values --session-id <SESSION> --sheet Sheet1 --range A1:C10 --values "[[1,2,3],[4,5,6]]"

# Get formulas
excelcli range get-formulas --session-id <SESSION> --sheet Sheet1 --range A1:D10

# Set formulas
excelcli range set-formulas --session-id <SESSION> --sheet Sheet1 --range A1 --formulas "[[=SUM(B1:B10)]]"

# Apply formatting
excelcli range format-range --session-id <SESSION> --sheet Sheet1 --range A1:E1 --bold --font-size 12 --h-align Center
excelcli range format-range --session-id <SESSION> --sheet Sheet1 --range D2:D100 --fill-color "#FFFF00"

# Set number format
excelcli range set-number-format --session-id <SESSION> --sheet Sheet1 --range D2:D100 --format "$#,##0.00"
excelcli range set-number-format --session-id <SESSION> --sheet Sheet1 --range E2:E100 --format "0.00%"

# Add data validation
excelcli range validate-range --session-id <SESSION> --sheet Sheet1 --range F2:F100 --type list --formula1 "Active,Inactive,Pending"

# Add hyperlink
excelcli range add-hyperlink --session-id <SESSION> --cell-address Sheet1!A1 --url "https://example.com" --display-text "Click Here"

# Merge cells
excelcli range merge-cells --session-id <SESSION> --sheet Sheet1 --range A1:D1
```

### Conditional Formatting

```bash
# Add conditional formatting rule (highlight cells > 100)
excelcli conditionalformat add-rule --session-id <SESSION> --sheet Sheet1 --range A1:A10 --rule-type cell-value --operator greater --formula1 100 --interior-color "#FFFF00"

# Add expression-based rule
excelcli conditionalformat add-rule --session-id <SESSION> --sheet Sheet1 --range B1:B10 --rule-type expression --formula1 "=B1>AVERAGE($B$1:$B$10)" --interior-color "#90EE90"

# Clear conditional formatting
excelcli conditionalformat clear-rules --session-id <SESSION> --sheet Sheet1 --range A1:A10
```

---

## ⚙️ System Requirements

| Requirement | Details | Why Required |
|-------------|---------|--------------|
| **Windows OS** | Windows 10/11 or Server 2016+ | COM interop is Windows-specific |
| **Microsoft Excel** | Excel 2016 or later | CLI controls actual Excel application |
| **.NET 10 Runtime** | [Download](https://dotnet.microsoft.com/download/dotnet/10.0) | Required to run .NET global tools |

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

```powershell
# PowerShell script example
$files = Get-ChildItem *.xlsx
foreach ($file in $files) {
    $session = excelcli session open $file.Name | Select-String "Session ID: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    excelcli powerquery refresh --session-id $session --query "Sales Data"
    excelcli datamodel refresh --session-id $session
    excelcli session save $session
    excelcli session close $session
}
```

### CI/CD Integration

```yaml
# GitHub Actions example
- name: Install ExcelMcp.CLI
  run: dotnet tool install --global Sbroenne.ExcelMcp.CLI

- name: Process Excel Files
  run: |
    SESSION=$(excelcli session open data.xlsx | grep "Session ID:" | cut -d' ' -f3)
    excelcli powerquery create --session-id $SESSION --query "Query1" --m-file queries/query1.pq
    excelcli powerquery refresh --session-id $SESSION --query "Query1"
    excelcli session save $SESSION
    excelcli session close $SESSION
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
