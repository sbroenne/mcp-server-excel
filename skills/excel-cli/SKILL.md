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
excelcli -q session open C:\Data\Report.xlsx
excelcli -q range set-values --session 1 --sheet Sheet1 --range A1:B2 --values-json '[["Name","Value"],["Test",123]]'
excelcli -q session close --session 1 --save
```

## Core Workflow

1. Open session: `excelcli -q session open <file>` → returns JSON with session ID
2. Run commands with `--session <id>`
3. Close and save: `excelcli -q session close --session <id> --save`

**Agent-friendly flags:**
- `-q` / `--quiet`: Suppress banner, output JSON only (recommended for agents)
- `--save`: Save changes before closing session
- Banner auto-suppresses when output is piped

## Check for Updates

```powershell
# Check if update available
excelcli -q version --check
# {"currentVersion":"1.2.0","latestVersion":"1.3.0","updateAvailable":true,...}

# Update to latest version
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

## Installation

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

## Commands

### Session Management

```powershell
excelcli session open C:\Data\Report.xlsx
excelcli session list
excelcli session close --session 1 --save
excelcli session close --session 1  # close without saving
excelcli create-empty C:\Data\New.xlsx --overwrite
```

### Range Operations

```powershell
# Value operations
excelcli range get-values --session 1 --sheet Sheet1 --range A1:D10
excelcli range set-values --session 1 --sheet Sheet1 --range A1 --values-json '[["Header1","Header2"]]'
excelcli range get-used-range --session 1 --sheet Sheet1
excelcli range get-current-region --session 1 --sheet Sheet1 --cell A1
excelcli range get-info --session 1 --sheet Sheet1 --range A1:D10

# Formula operations
excelcli range get-formulas --session 1 --sheet Sheet1 --range B2:B10
excelcli range set-formulas --session 1 --sheet Sheet1 --range B2 --formulas-json '[["=SUM(A1:A10)"]]'

# Clear operations
excelcli range clear --session 1 --sheet Sheet1 --range A1:D10  # clears all (values, formulas, formats)
excelcli range clear-contents --session 1 --sheet Sheet1 --range A1:D10  # preserves formatting
excelcli range clear-formats --session 1 --sheet Sheet1 --range A1:D10  # preserves values

# Copy operations
excelcli range copy --session 1 --source-sheet Sheet1 --source-range A1:B5 --target-sheet Sheet2 --target-range A1
excelcli range copy-values --session 1 --source-sheet Sheet1 --source-range A1:B5 --target-sheet Sheet2 --target-range A1
excelcli range copy-formulas --session 1 --source-sheet Sheet1 --source-range A1:B5 --target-sheet Sheet2 --target-range A1

# Insert/Delete operations
excelcli range insert-cells --session 1 --sheet Sheet1 --range A1:B5 --shift-direction down
excelcli range delete-cells --session 1 --sheet Sheet1 --range A1:B5 --shift-direction up
excelcli range insert-rows --session 1 --sheet Sheet1 --range A1:A5
excelcli range delete-rows --session 1 --sheet Sheet1 --range A1:A5
excelcli range insert-columns --session 1 --sheet Sheet1 --range A1:C1
excelcli range delete-columns --session 1 --sheet Sheet1 --range A1:C1

# Find/Replace operations
excelcli range find --session 1 --sheet Sheet1 --range A1:Z100 --search-value "test" --match-case false
excelcli range replace --session 1 --sheet Sheet1 --range A1:Z100 --find-value "old" --replace-value "new" --replace-all true

# Sort operations
excelcli range sort --session 1 --sheet Sheet1 --range A1:D10 --sort-columns-json '[{"columnIndex":1,"ascending":true}]' --has-headers true
```

### Number Formatting

```powershell
excelcli range get-number-formats --session 1 --sheet Sheet1 --range B2:B10
excelcli range set-number-format --session 1 --sheet Sheet1 --range B2:B10 --format-code "$#,##0.00"
excelcli range set-number-formats --session 1 --sheet Sheet1 --range B2:C10 --formats-json '[["$#,##0.00","0.00%"],...]'
```

### Cell Styling

```powershell
# Built-in styles (recommended - theme-aware)
excelcli range set-style --session 1 --sheet Sheet1 --range A1 --style-name "Heading 1"
excelcli range get-style --session 1 --sheet Sheet1 --range A1

# Custom formatting (for specific needs)
excelcli range format-range --session 1 --sheet Sheet1 --range A1:D1 \
  --font-name "Arial" --font-size 14 --bold true --font-color "#0000FF" \
  --fill-color "#FFFF00" --horizontal-alignment center --wrap-text true

# AutoFit
excelcli range auto-fit-columns --session 1 --sheet Sheet1 --range A:D
excelcli range auto-fit-rows --session 1 --sheet Sheet1 --range 1:10
```

Built-in styles: `Normal`, `Heading 1-4`, `Title`, `Total`, `Input`, `Output`, `Calculation`, `Good`, `Bad`, `Neutral`, `Accent1-6`, `Currency`, `Percent`, `Comma`

### Data Validation

```powershell
# List validation (dropdown)
excelcli range set-validation --session 1 --sheet Sheet1 --range B2:B100 \
  --validation-type list --formula1 "Yes,No,Maybe" --show-dropdown true

# Number range validation
excelcli range set-validation --session 1 --sheet Sheet1 --range C2:C100 \
  --validation-type whole --validation-operator between --formula1 1 --formula2 100 \
  --show-error-alert true --error-title "Invalid" --error-message "Enter 1-100"

# Get/Remove validation
excelcli range get-validation --session 1 --sheet Sheet1 --range B2
excelcli range clear-validation --session 1 --sheet Sheet1 --range B2:B100
```

Validation types: `whole`, `decimal`, `list`, `date`, `time`, `textLength`, `custom`
Operators: `between`, `notBetween`, `equal`, `notEqual`, `greaterThan`, `lessThan`, `greaterThanOrEqual`, `lessThanOrEqual`

### Hyperlinks

```powershell
excelcli range set-hyperlink --session 1 --sheet Sheet1 --cell A1 --url "https://example.com" --display-text "Click here" --tooltip "Visit site"
excelcli range get-hyperlink --session 1 --sheet Sheet1 --cell A1
excelcli range get-hyperlinks --session 1 --sheet Sheet1
excelcli range remove-hyperlink --session 1 --sheet Sheet1 --range A1:A10
```

### Merge/Unmerge Cells

```powershell
excelcli range merge-cells --session 1 --sheet Sheet1 --range A1:D1
excelcli range unmerge-cells --session 1 --sheet Sheet1 --range A1:D1
excelcli range get-merge-info --session 1 --sheet Sheet1 --range A1:D10
```

### Cell Protection

```powershell
# Lock/unlock cells (requires worksheet protection to take effect)
excelcli range set-cell-lock --session 1 --sheet Sheet1 --range A1:D10 --locked true
excelcli range get-cell-lock --session 1 --sheet Sheet1 --range A1
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

# Cross-file operations (atomic - no session needed)
excelcli sheet copy-to-file --source-file "C:\Data\Source.xlsx" --source-sheet "Data" \
  --target-file "C:\Data\Target.xlsx" --target-sheet-name "ImportedData"
excelcli sheet move-to-file --source-file "C:\Data\Source.xlsx" --source-sheet "Archive" \
  --target-file "C:\Data\Archive.xlsx"

# Tab color
excelcli sheet set-tab-color --session 1 --sheet Sheet1 --red 255 --green 0 --blue 0
excelcli sheet get-tab-color --session 1 --sheet Sheet1
excelcli sheet clear-tab-color --session 1 --sheet Sheet1

# Visibility
excelcli sheet set-visibility --session 1 --sheet Hidden --visibility hidden  # visible, hidden, veryhidden
excelcli sheet get-visibility --session 1 --sheet Hidden
excelcli sheet show --session 1 --sheet Hidden
excelcli sheet hide --session 1 --sheet Temp
excelcli sheet very-hide --session 1 --sheet Secret  # requires code to unhide
```

### Tables

```powershell
excelcli table list --session 1
excelcli table create --session 1 --sheet Sheet1 --range A1:D10 --name "SalesData" --has-headers
excelcli table get --session 1 --name SalesData
excelcli table resize --session 1 --name SalesData --range A1:E20
excelcli table add-column --session 1 --name SalesData --column-name "Total"
excelcli table delete-column --session 1 --name SalesData --column-name "Temp"
excelcli table filter --session 1 --name SalesData --column-name "Region" --filter-values "North,South"
excelcli table clear-filters --session 1 --name SalesData
excelcli table sort --session 1 --name SalesData --column-name "Amount" --descending
excelcli table set-total-row --session 1 --name SalesData --show-totals --totals-json '{"Amount":"sum","Count":"count"}'
excelcli table add-to-datamodel --session 1 --name SalesData
excelcli table rename --session 1 --name SalesData --new-name "Sales"
excelcli table delete --session 1 --name OldTable

# Structured references (for formulas)
excelcli table get-structured-reference --session 1 --name SalesData --region data  # all, data, headers, totals
excelcli table get-structured-reference --session 1 --name SalesData --region data --column-name "Amount"
```

### Named Ranges

```powershell
excelcli namedrange list --session 1
excelcli namedrange create --session 1 --name "TaxRate" --refers-to "=0.25"
excelcli namedrange create --session 1 --name "DataRange" --refers-to "=Sheet1!$A$1:$D$100"
excelcli namedrange update --session 1 --name "TaxRate" --refers-to "=0.30"
excelcli namedrange delete --session 1 --name "OldName"
```

### Power Query

```powershell
excelcli powerquery list --session 1
excelcli powerquery get --session 1 --name "SalesQuery"
excelcli powerquery create --session 1 --name "CsvImport" --m-code 'let Source = Csv.Document(File.Contents("C:\Data\sales.csv")) in Source' --load-destination worksheet
excelcli powerquery update --session 1 --name "CsvImport" --m-code 'let Source = Csv.Document(File.Contents("C:\Data\sales_new.csv")) in Source'
excelcli powerquery refresh --session 1 --name "CsvImport" --timeout 120
excelcli powerquery delete --session 1 --name "OldQuery"
```

Load destinations: `worksheet`, `data-model`, `both`, `connection-only`

### Data Model (Power Pivot)

```powershell
excelcli datamodel list-tables --session 1
excelcli datamodel list-measures --session 1
excelcli datamodel create-measure --session 1 --table "Sales" --name "TotalRevenue" --dax "SUM(Sales[Amount])"
excelcli datamodel update-measure --session 1 --table "Sales" --name "TotalRevenue" --dax "SUMX(Sales, Sales[Qty] * Sales[Price])"
excelcli datamodel delete-measure --session 1 --table "Sales" --name "OldMeasure"
excelcli datamodel list-relationships --session 1
excelcli datamodel create-relationship --session 1 --from-table "Sales" --from-column "ProductID" --to-table "Products" --to-column "ID"
```

### PivotTables

```powershell
# Lifecycle
excelcli pivottable list --session 1
excelcli pivottable create-from-table --session 1 --source-table "SalesData" --target-sheet "PivotSheet" --target-cell A1 --name "SalesPivot"
excelcli pivottable create-from-datamodel --session 1 --target-sheet "Analysis" --target-cell A1 --name "ModelPivot"
excelcli pivottable refresh --session 1 --name "SalesPivot"
excelcli pivottable delete --session 1 --name "OldPivot"

# Fields
excelcli pivottable add-field --session 1 --name "SalesPivot" --field "Region" --area row
excelcli pivottable add-field --session 1 --name "SalesPivot" --field "Amount" --area data --function sum
excelcli pivottable remove-field --session 1 --name "SalesPivot" --field "Category"
excelcli pivottable set-field-function --session 1 --name "SalesPivot" --field "Amount" --aggregation-function average
excelcli pivottable set-field-name --session 1 --name "SalesPivot" --field "Sum of Amount" --custom-name "Total Revenue"
excelcli pivottable set-field-format --session 1 --name "SalesPivot" --field "Amount" --number-format "$#,##0.00"

# Grouping
excelcli pivottable group-by-date --session 1 --name "SalesPivot" --field "OrderDate" --interval months  # days, months, quarters, years
excelcli pivottable group-by-numeric --session 1 --name "SalesPivot" --field "Price" --start 0 --end 1000 --interval-size 100

# Calculated fields (non-OLAP)
excelcli pivottable create-calculated-field --session 1 --name "SalesPivot" --field "Profit" --formula "=Sales-Cost"
excelcli pivottable list-calculated-fields --session 1 --name "SalesPivot"
excelcli pivottable delete-calculated-field --session 1 --name "SalesPivot" --field "Profit"

# Calculated members (OLAP/Data Model)
excelcli pivottable list-calculated-members --session 1 --name "ModelPivot"
excelcli pivottable create-calculated-member --session 1 --name "ModelPivot" --member-name "YTDSales" --formula "..." --member-type measure
excelcli pivottable delete-calculated-member --session 1 --name "ModelPivot" --member-name "YTDSales"
```

Field areas: `row`, `column`, `data`, `page`
Functions: `sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevP`, `var`, `varP`
Grouping intervals: `days`, `months`, `quarters`, `years`

### Slicers

```powershell
# PivotTable slicers
excelcli slicer create-slicer --session 1 --pivot-name "SalesPivot" --field-name "Region" --destination-sheet "Dashboard" --position E1
excelcli slicer list-slicers --session 1
excelcli slicer list-slicers --session 1 --pivot-name "SalesPivot"
excelcli slicer set-slicer-selection --session 1 --slicer-name "RegionSlicer" --selected-items '["North","South"]'
excelcli slicer delete-slicer --session 1 --slicer-name "RegionSlicer"

# Table slicers
excelcli slicer create-table-slicer --session 1 --table-name "SalesData" --column-name "Category" --destination-sheet "Dashboard" --position G1
excelcli slicer list-table-slicers --session 1
excelcli slicer set-table-slicer-selection --session 1 --slicer-name "CategorySlicer" --selected-items "Electronics,Furniture"
excelcli slicer delete-table-slicer --session 1 --slicer-name "CategorySlicer"
```

Note: `--selected-items` accepts JSON array or comma-separated values. Use `--clear-first false` to add to existing selection.

### Charts

```powershell
# Lifecycle
excelcli chart list --session 1
excelcli chart create-from-range --session 1 --sheet Sheet1 --source-range A1:B10 --chart-type 51 --left 100 --top 50 --name "SalesChart"
excelcli chart create-from-pivottable --session 1 --pivot-name "SalesPivot" --sheet Dashboard --chart-type 5 --left 100 --top 50
excelcli chart delete --session 1 --chart-name "OldChart"

# Configuration
excelcli chart set-title --session 1 --chart-name "SalesChart" --title "Monthly Sales"
excelcli chart set-axis-title --session 1 --chart-name "SalesChart" --axis-type Value --title "Revenue ($)"
excelcli chart set-chart-type --session 1 --chart-name "SalesChart" --chart-type 4
excelcli chart move --session 1 --chart-name "SalesChart" --left 100 --top 50 --width 400 --height 300
excelcli chart set-placement --session 1 --chart-name "SalesChart" --placement 2
excelcli chart fit-to-range --session 1 --chart-name "SalesChart" --sheet Sheet1 --source-range D1:H20

# Data labels and formatting
excelcli chart set-data-labels --session 1 --chart-name "SalesChart" --show-value true --label-position OutsideEnd
excelcli chart set-style --session 1 --chart-name "SalesChart" --style-id 5
excelcli chart show-legend --session 1 --chart-name "SalesChart" --visible true --legend-position Bottom

# Axis configuration
excelcli chart set-axis-scale --session 1 --chart-name "SalesChart" --axis-type Value --minimum-scale 0 --maximum-scale 1000
excelcli chart set-axis-number-format --session 1 --chart-name "SalesChart" --axis-type Value --number-format "$#,##0"
excelcli chart set-gridlines --session 1 --chart-name "SalesChart" --axis-type Value --show-major true --show-minor false

# Series and trendlines
excelcli chart add-series --session 1 --chart-name "SalesChart" --series-name "Q2" --values-range "Sheet1!C2:C10"
excelcli chart set-series-format --session 1 --chart-name "SalesChart" --series-index 1 --marker-style Circle --marker-size 8
excelcli chart add-trendline --session 1 --chart-name "SalesChart" --series-index 1 --trendline-type Linear --display-equation true
excelcli chart list-trendlines --session 1 --chart-name "SalesChart" --series-index 1
```

Chart types: Use Excel XlChartType constants (e.g., 51=ColumnClustered, 4=Line, 5=Pie)
Placement: 1=Move and size with cells, 2=Move only, 3=Don't move or size

### Conditional Formatting

```powershell
excelcli conditionalformat add-color-scale --session 1 --sheet Sheet1 --range B2:B100 --min-color "#FF0000" --max-color "#00FF00"
excelcli conditionalformat add-data-bar --session 1 --sheet Sheet1 --range C2:C100 --bar-color "#0000FF"
excelcli conditionalformat add-icon-set --session 1 --sheet Sheet1 --range D2:D100 --icon-set "3Arrows"
excelcli conditionalformat add-cell-value --session 1 --sheet Sheet1 --range E2:E100 --operator greaterThan --value 100 --format-fill "#FFFF00"
excelcli conditionalformat clear --session 1 --sheet Sheet1 --range B2:E100
```

### Connections

```powershell
excelcli connection list --session 1
excelcli connection get --session 1 --name "SQLServer_Sales"
excelcli connection refresh --session 1 --name "SQLServer_Sales"
excelcli connection refresh-all --session 1
```

### VBA (requires Trust Center enabled)

```powershell
excelcli vba list --session 1
excelcli vba get --session 1 --module "Module1"
excelcli vba create --session 1 --module "Utilities" --code "Sub Hello()\n    MsgBox \"Hello!\"\nEnd Sub"
excelcli vba update --session 1 --module "Utilities" --code "Sub Hello()\n    MsgBox \"Updated!\"\nEnd Sub"
excelcli vba delete --session 1 --module "OldModule"
excelcli vba run --session 1 --macro "Utilities.Hello"
excelcli vba export --session 1 --module "Module1" --file "C:\Backup\Module1.bas"
excelcli vba import --session 1 --file "C:\Code\NewModule.bas"
```

## Example Workflows

### Import CSV and Create Dashboard

```powershell
# Create workbook
excelcli -q create-empty C:\Reports\Dashboard.xlsx

# Open session
excelcli -q session open C:\Reports\Dashboard.xlsx
# Returns: {"success":true,"sessionId":1,...}

# Import CSV via Power Query
excelcli -q powerquery create --session 1 --name "SalesData" `
  --m-code 'let Source = Csv.Document(File.Contents("C:\Data\sales.csv"),[Delimiter=",",Encoding=65001]) in Source' `
  --load-destination data-model

# Refresh to load data
excelcli -q powerquery refresh --session 1 --name "SalesData" --timeout 120

# Create DAX measure
excelcli -q datamodel create-measure --session 1 --table "SalesData" --name "TotalSales" --dax "SUM(SalesData[Amount])"

# Create PivotTable
excelcli -q pivottable create-from-datamodel --session 1 --target-sheet "Analysis" --target-cell A1 --name "SalesPivot"
excelcli -q pivottable add-field --session 1 --name "SalesPivot" --field "Region" --area row
excelcli -q pivottable add-field --session 1 --name "SalesPivot" --field "TotalSales" --area data

# Save and close
excelcli -q session close --session 1 --save
```

### Batch Update Multiple Files

```powershell
Get-ChildItem C:\Data\*.xlsx | ForEach-Object {
  excelcli -q session open $_.FullName
  excelcli -q range set-values --session 1 --sheet Summary --range A1 --values-json '[["Updated: 2026-01-28"]]'
  excelcli -q session close --session 1 --save
}
```

### Format Financial Report

```powershell
excelcli -q session open C:\Reports\Financial.xlsx

# Format currency column
excelcli -q range set-number-format --session 1 --sheet Data --range C2:C100 --format-code "$#,##0.00"

# Format percentages
excelcli -q range set-number-format --session 1 --sheet Data --range D2:D100 --format-code "0.00%"

# Add conditional formatting for negative values
excelcli -q conditionalformat add-cell-value --session 1 --sheet Data --range C2:C100 --operator lessThan --value 0 --format-font-color "#FF0000"

# Header styling
excelcli -q range set-font --session 1 --sheet Data --range A1:E1 --bold --fill-color "#4472C4" --font-color "#FFFFFF"

excelcli -q session close --session 1 --save
```

## Tips

- **Use `-q` flag**: Always use `-q` for clean JSON output (no banner)
- **Check help**: `excelcli <command> --help` shows all options
- **JSON values**: Use `--values-json` for 2D arrays: `'[["A","B"],["C","D"]]'`
- **Session reuse**: Keep session open for multiple operations (faster than open/close each time)
- **Timeout**: Power Query refresh may need `--timeout 120` or higher for large datasets
- **US format codes**: Use US locale for number formats (`#,##0.00` not `#.##0,00`)

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
