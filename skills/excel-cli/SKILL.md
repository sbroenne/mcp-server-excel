---
name: excel-cli
description: >
  Automate Microsoft Excel on Windows via CLI. Use when creating, reading, 
  or modifying Excel workbooks from scripts, CI/CD, or coding agents.
  Supports Power Query, DAX, PivotTables, Tables, Ranges, Charts, VBA.
  Triggers: Excel, spreadsheet, workbook, xlsx, excelcli, CLI automation.
allowed-tools: Cmd(excelcli:*),PowerShell(excelcli:*)
disable-model-invocation: true
license: MIT
version: 1.3.0
tags:
  - excel
  - cli
  - automation
  - windows
  - powerquery
  - dax
  - scripting
repository: https://github.com/sbroenne/mcp-server-excel
documentation: https://excelmcpserver.dev/
---

# Excel Automation with excelcli

## Quick Start

```powershell
# Create new workbook
excelcli -q session create C:\Data\Report.xlsx
excelcli -q range set-values --session 1 --sheet Sheet1 --range A1:B2 --values '[["Name","Value"],["Test",123]]'
excelcli -q session close --session 1 --save

# Open existing workbook
excelcli -q session open C:\Data\Report.xlsx
excelcli -q range get-values --session 1 --sheet Sheet1 --range A1:B2
excelcli -q session close --session 1
```

## Core Workflow

1. Create or open session:
   - New file: `excelcli -q session create <file>` → creates file and returns session ID (optimized single Excel startup)
   - Existing file: `excelcli -q session open <file>` → returns session ID
2. Run commands with `--session <id>`
3. Close and save: `excelcli -q session close --session <id> --save`

**Session Timeout:**
- Default: 5 minutes per operation (prevents hangs when Excel gets stuck)
- Custom timeout: `excelcli -q session open <file> --timeout 600` (10 minutes)
- Range: 10-3600 seconds

**Agent-friendly flags:**
- `-q` / `--quiet`: Suppress banner, output JSON only (recommended for agents)
- `--save`: Save changes before closing session
- Banner auto-suppresses when output is piped

## Check for Updates

```powershell
# Check if update available
excelcli --version
# (shows current version and checks against latest NuGet version)

# Update to latest version
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

## Installation

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

## Commands

### Discover Available Actions

```powershell
# List actions for all commands
excelcli actions

# List actions for a single command
excelcli actions range
```

### Session Management

```powershell
excelcli session create C:\Data\New.xlsx  # create new workbook, returns session ID
excelcli session open C:\Data\Report.xlsx  # open existing workbook
excelcli session open C:\Data\Large.xlsx --timeout 600  # 10-minute timeout for heavy files
excelcli session list
excelcli session close --session 1 --save
excelcli session close --session 1  # close without saving
```

### Range Operations

```powershell
# Value operations
excelcli range get-values --session 1 --sheet Sheet1 --range A1:D10
excelcli range set-values --session 1 --sheet Sheet1 --range A1 --values '[["Header1","Header2"]]'
excelcli range get-used-range --session 1 --sheet Sheet1
excelcli range get-current-region --session 1 --sheet Sheet1 --range A1
excelcli range get-info --session 1 --sheet Sheet1 --range A1:D10

# Formula operations
excelcli range get-formulas --session 1 --sheet Sheet1 --range B2:B10

# Clear operations
excelcli range clear-all --session 1 --sheet Sheet1 --range A1:D10  # clears all (values, formulas, formats)
excelcli range clear-contents --session 1 --sheet Sheet1 --range A1:D10  # preserves formatting
excelcli range clear-formats --session 1 --sheet Sheet1 --range A1:D10  # preserves values

# Insert/Delete operations
excelcli range insert-cells --session 1 --sheet Sheet1 --range A1:B5
excelcli range delete-cells --session 1 --sheet Sheet1 --range A1:B5
excelcli range insert-rows --session 1 --sheet Sheet1 --range A1:A5
excelcli range delete-rows --session 1 --sheet Sheet1 --range A1:A5
excelcli range insert-columns --session 1 --sheet Sheet1 --range A1:C1
excelcli range delete-columns --session 1 --sheet Sheet1 --range A1:C1
```

Note: `excelcli range` currently exposes `--sheet`, `--range`, and `--values`. Use `excelcli actions range` and `excelcli range --help` for the full action list and available flags.

### Number Formatting

```powershell
excelcli range get-number-formats --session 1 --sheet Sheet1 --range B2:B10
```

### Merge/Unmerge Cells

```powershell
excelcli range merge-cells --session 1 --sheet Sheet1 --range A1:D1
excelcli range unmerge-cells --session 1 --sheet Sheet1 --range A1:D1
excelcli range get-merge-info --session 1 --sheet Sheet1 --range A1:D10
```

### Worksheets

```powershell
# Basic operations
excelcli sheet list --session 1
excelcli sheet create --session 1 --name "NewSheet"
excelcli sheet rename --session 1 --sheet Sheet1 --new-name "DataSheet"
excelcli sheet copy --session 1 --source-sheet Sheet1 --target-sheet "Sheet1_Copy"
excelcli sheet delete --session 1 --sheet OldSheet
excelcli sheet move --session 1 --sheet Sheet1 --before-sheet Sheet2  # or --after-sheet
```

### Tables

```powershell
excelcli table list --session 1
excelcli table create --session 1 --sheet Sheet1 --range A1:D10 --table "SalesData" --has-headers
excelcli table read --session 1 --table SalesData
excelcli table resize --session 1 --table SalesData --range A1:E20
excelcli table set-style --session 1 --table SalesData --style "TableStyleMedium2"
excelcli table toggle-totals --session 1 --table SalesData --has-headers
excelcli table set-column-total --session 1 --table SalesData --new-name "Amount" --style "sum"
excelcli table append --session 1 --table SalesData --csv-data "Region,Amount\nNorth,100\nSouth,200"
excelcli table get-data --session 1 --table SalesData --visible-only
excelcli table add-to-datamodel --session 1 --table SalesData
excelcli table rename --session 1 --table SalesData --new-name "Sales"
excelcli table delete --session 1 --table OldTable
```

### Named Ranges

```powershell
excelcli namedrange list --session 1
excelcli namedrange read --session 1 --name "TaxRate"
excelcli namedrange write --session 1 --name "TaxRate" --value 0.25
excelcli namedrange create --session 1 --name "TaxRate" --refers-to "=0.25"
excelcli namedrange create --session 1 --name "DataRange" --refers-to "=Sheet1!$A$1:$D$100"
excelcli namedrange update --session 1 --name "TaxRate" --refers-to "=0.30"
excelcli namedrange delete --session 1 --name "OldName"
```

### Power Query

```powershell
excelcli powerquery list --session 1
excelcli powerquery view --session 1 --query "SalesQuery"
excelcli powerquery create --session 1 --query "CsvImport" --mcode 'let Source = Csv.Document(File.Contents("C:\Data\sales.csv")) in Source' --load-destination worksheet
excelcli powerquery update --session 1 --query "CsvImport" --mcode 'let Source = Csv.Document(File.Contents("C:\Data\sales_new.csv")) in Source'
excelcli powerquery refresh --session 1 --query "CsvImport"
excelcli powerquery delete --session 1 --query "OldQuery"
excelcli powerquery refresh-all --session 1
excelcli powerquery get-load-config --session 1 --query "CsvImport"
```

Load destinations: `worksheet`, `data-model`, `both`, `connection-only`

### Data Model (Power Pivot)

```powershell
excelcli datamodel list-tables --session 1
excelcli datamodel read-table --session 1 --table "Sales" --max-rows 100
excelcli datamodel list-columns --session 1 --table "Sales"
excelcli datamodel list-measures --session 1
excelcli datamodel create-measure --session 1 --table "Sales" --measure "TotalRevenue" --expression "SUM(Sales[Amount])"
excelcli datamodel update-measure --session 1 --table "Sales" --measure "TotalRevenue" --expression "SUMX(Sales, Sales[Qty] * Sales[Price])"
excelcli datamodel delete-measure --session 1 --table "Sales" --measure "OldMeasure"
excelcli datamodel read --session 1 --table "Sales" --measure "TotalRevenue"
excelcli datamodel rename-table --session 1 --table "Sales" --new-name "SalesFact"
excelcli datamodel delete-table --session 1 --table "SalesFact"
excelcli datamodel read-info --session 1
excelcli datamodel refresh --session 1
excelcli datamodel evaluate --session 1 --dax-query "EVALUATE Sales"
excelcli datamodel execute-dmv --session 1 --dmv-query "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES"
```

### PivotTables

```powershell
# Lifecycle
excelcli pivottable list --session 1
excelcli pivottable read --session 1 --pivot-table "SalesPivot"
excelcli pivottable create-from-table --session 1 --table "SalesData" --dest-sheet "PivotSheet" --dest-cell A1 --pivot-table "SalesPivot"
excelcli pivottable create-from-datamodel --session 1 --table "Sales" --dest-sheet "Analysis" --dest-cell A1 --pivot-table "ModelPivot"
excelcli pivottable refresh --session 1 --pivot-table "SalesPivot"
excelcli pivottable delete --session 1 --pivot-table "OldPivot"
```

### Slicers

```powershell
# PivotTable slicers
excelcli slicer create-slicer --session 1 --pivottable "SalesPivot" --source-field "Region" --destination-sheet "Dashboard" --left 100 --top 20
excelcli slicer list-slicers --session 1
excelcli slicer list-slicers --session 1 --pivottable "SalesPivot"
excelcli slicer set-slicer-selection --session 1 --slicer "RegionSlicer" --selected-items '["North","South"]'
excelcli slicer delete-slicer --session 1 --slicer "RegionSlicer"

# Table slicers
excelcli slicer create-table-slicer --session 1 --table "SalesData" --column "Category" --destination-sheet "Dashboard" --left 300 --top 20
excelcli slicer list-table-slicers --session 1
excelcli slicer set-table-slicer-selection --session 1 --slicer "CategorySlicer" --selected-items "Electronics,Furniture"
excelcli slicer delete-table-slicer --session 1 --slicer "CategorySlicer"
```

Note: `--selected-items` accepts JSON array or comma-separated values. Use `--multi-select` to add to existing selection.

### Charts

```powershell
# Lifecycle
excelcli chart list --session 1 --sheet Sheet1
excelcli chart create-from-range --session 1 --sheet Sheet1 --source-range A1:B10 --chart-type 51 --chart "SalesChart"
excelcli chart create-from-pivottable --session 1 --pivot-table "SalesPivot" --sheet Dashboard --chart-type 5 --chart "PivotChart"
excelcli chart read --session 1 --sheet Sheet1 --chart "SalesChart"
excelcli chart move --session 1 --chart "SalesChart"
excelcli chart fit-to-range --session 1 --chart "SalesChart" --sheet Sheet1 --target-range D1:H20
excelcli chart delete --session 1 --sheet Sheet1 --chart "OldChart"

# Configuration
excelcli chartconfig set-title --session 1 --chart "SalesChart" --title "Monthly Sales"
excelcli chartconfig set-axis-title --session 1 --chart "SalesChart" --axis Value --title "Revenue ($)"
excelcli chartconfig set-chart-type --session 1 --chart "SalesChart" --chart-type 4
excelcli chartconfig set-placement --session 1 --chart "SalesChart" --placement 2

# Data labels and formatting
excelcli chartconfig set-data-labels --session 1 --chart "SalesChart" --show-value --label-position OutsideEnd
excelcli chartconfig set-style --session 1 --chart "SalesChart" --style-id 5
excelcli chartconfig show-legend --session 1 --chart "SalesChart" --visible --legend-position Bottom

# Axis configuration
excelcli chartconfig set-axis-scale --session 1 --chart "SalesChart" --axis Value --minimum-scale 0 --maximum-scale 1000
excelcli chartconfig set-axis-number-format --session 1 --chart "SalesChart" --axis Value --number-format "$#,##0"
excelcli chartconfig set-gridlines --session 1 --chart "SalesChart" --axis Value --show-major --show-minor

# Series and trendlines
excelcli chartconfig add-series --session 1 --chart "SalesChart" --series-name "Q2" --values-range "Sheet1!C2:C10"
excelcli chartconfig set-series-format --session 1 --chart "SalesChart" --series-index 1 --marker-style Circle --marker-size 8
excelcli chartconfig add-trendline --session 1 --chart "SalesChart" --series-index 1 --trendline-type Linear --display-equation
excelcli chartconfig list-trendlines --session 1 --chart "SalesChart" --series-index 1
```

Chart types: Use Excel XlChartType constants (e.g., 51=ColumnClustered, 4=Line, 5=Pie)
Placement: 1=Move and size with cells, 2=Move only, 3=Don't move or size

### Conditional Formatting

```powershell
excelcli conditionalformat add-rule --session 1 --sheet Sheet1 --range B2:B100 --rule-type ColorScale --format-style "RedGreen"
excelcli conditionalformat add-rule --session 1 --sheet Sheet1 --range E2:E100 --rule-type CellValue --formula ">100" --format-style "YellowFill"
excelcli conditionalformat clear-rules --session 1 --sheet Sheet1 --range B2:E100
```

### Connections

```powershell
excelcli connection list --session 1
excelcli connection view --session 1 --connection "SQLServer_Sales"
excelcli connection test --session 1 --connection "SQLServer_Sales"
excelcli connection refresh --session 1 --connection "SQLServer_Sales"
excelcli connection get-properties --session 1 --connection "SQLServer_Sales"
excelcli connection set-properties --session 1 --connection "SQLServer_Sales" --refresh-on-open --enable-refresh
```

### VBA (requires Trust Center enabled)

```powershell
excelcli vba list --session 1
excelcli vba view --session 1 --module "Module1"
excelcli vba import --session 1 --module "Utilities" --code "Sub Hello()\n    MsgBox \"Hello!\"\nEnd Sub"
excelcli vba update --session 1 --module "Utilities" --code "Sub Hello()\n    MsgBox \"Updated!\"\nEnd Sub"
excelcli vba delete --session 1 --module "OldModule"
excelcli vba run --session 1 --macro "Utilities.Hello"
```

## Example Workflows

### Import CSV and Create Dashboard

```powershell
# Create workbook and get session ID
excelcli -q session create C:\Reports\Dashboard.xlsx
# Returns: {"success":true,"sessionId":1,...}

# Import CSV via Power Query
excelcli -q powerquery create --session 1 --query "SalesData" `
  --mcode 'let Source = Csv.Document(File.Contents("C:\Data\sales.csv"),[Delimiter=",",Encoding=65001]) in Source' `
  --load-destination data-model

# Refresh to load data
excelcli -q powerquery refresh --session 1 --query "SalesData"

# Create DAX measure
excelcli -q datamodel create-measure --session 1 --table "SalesData" --measure "TotalSales" --expression "SUM(SalesData[Amount])"

# Create PivotTable
excelcli -q pivottable create-from-datamodel --session 1 --table "SalesData" --dest-sheet "Analysis" --dest-cell A1 --pivot-table "SalesPivot"

# Save and close
excelcli -q session close --session 1 --save
```

### Batch Update Multiple Files

```powershell
Get-ChildItem C:\Data\*.xlsx | ForEach-Object {
  excelcli -q session open $_.FullName
  excelcli -q range set-values --session 1 --sheet Summary --range A1 --values '[["Updated: 2026-01-28"]]'
  excelcli -q session close --session 1 --save
}
```

### Format Financial Report

```powershell
excelcli -q session open C:\Reports\Financial.xlsx

# Format currency column
# Apply a built-in table style
excelcli -q table set-style --session 1 --table "DataTable" --style "TableStyleMedium2"

# Add conditional formatting for negative values
excelcli -q conditionalformat add-rule --session 1 --sheet Data --range C2:C100 --rule-type CellValue --formula "<0" --format-style "RedFill"

excelcli -q session close --session 1 --save
```

## Tips

- **Use `-q` flag**: Always use `-q` for clean JSON output (no banner)
- **Check help**: `excelcli <command> --help` shows all options
- **Session reuse**: Keep session open for multiple operations (faster than open/close each time)

### JSON Values Parameter

When using `--values` for range operations, pass a 2D JSON array. Use single quotes around the JSON in PowerShell:

```powershell
# Correct: single quotes preserve the JSON
excelcli -q range set-values --session 1 --sheet Sheet1 --range A1:B2 --values '[["Name","Value"],["Test",123]]'

# Multiple rows and columns
excelcli -q range set-values --session 1 --sheet Sheet1 --range A1:C4 --values '[["Product","Q1","Q2"],["Widget",100,150],["Gadget",80,90],["Device",200,180]]'
```

**Important**: The `--values` parameter expects a JSON 2D array where:
- Outer array = rows
- Inner arrays = cells in each row
- Strings must be quoted: `"text"`
- Numbers are unquoted: `123`, `45.67`
- The range dimensions should match the array dimensions

## Requirements

- Windows with Microsoft Excel 2016+
- .NET 10 Runtime
- VBA operations require "Trust access to VBA project object model" enabled

## Reference Documentation

See `references/` for detailed behavioral guidance and quirks:

- @references/behavioral-rules.md - Core execution rules (format cells, use Tables, etc.)
- @references/anti-patterns.md - Common mistakes to avoid
- @references/workflows.md - Production workflow patterns
- @references/excel_powerquery.md - Power Query quirks (timeout required, create vs update)
- @references/excel_datamodel.md - Data Model/DAX specifics (prerequisites, timeouts)
- @references/excel_table.md - Table operations and Data Model integration
- @references/excel_range.md - Range operations and number formatting
- @references/excel_worksheet.md - Worksheet operations
- @references/excel_chart.md - Chart positioning and configuration
- @references/excel_slicer.md - Slicer operations for PivotTables and Tables
- @references/excel_conditionalformat.md - Conditional formatting rules

````
