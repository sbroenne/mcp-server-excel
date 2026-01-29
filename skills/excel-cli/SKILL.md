---
name: excel-cli
description: >
  Automate Microsoft Excel on Windows via CLI. Use when creating, reading, 
  or modifying Excel workbooks from scripts, CI/CD, or coding agents.
  Supports Power Query, DAX, PivotTables, Tables, Ranges, Charts, VBA.
  Triggers: Excel, spreadsheet, workbook, xlsx, excelcli, CLI automation.
allowed-tools: Bash(excelcli:*)
disable-model-invocation: true
license: MIT
version: 1.2.0
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

```bash
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

```bash
# Check if update available
excelcli -q version --check
# {"currentVersion":"1.2.0","latestVersion":"1.3.0","updateAvailable":true,...}

# Update to latest version
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

## Installation

```bash
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

## Commands

### Session Management

```bash
excelcli session open C:\Data\Report.xlsx
excelcli session list
excelcli session close --session 1 --save
excelcli session close --session 1  # close without saving
excelcli create-empty C:\Data\New.xlsx --overwrite
```

### Range Operations

```bash
excelcli range get-values --session 1 --sheet Sheet1 --range A1:D10
excelcli range set-values --session 1 --sheet Sheet1 --range A1 --values-json '[["Header1","Header2"]]'
excelcli range set-formulas --session 1 --sheet Sheet1 --range B2 --formulas-json '[["=SUM(A1:A10)"]]'
excelcli range get-used-range --session 1 --sheet Sheet1
excelcli range clear --session 1 --sheet Sheet1 --range A1:D10
excelcli range copy --session 1 --source-sheet Sheet1 --source-range A1:B5 --target-sheet Sheet2 --target-range A1
excelcli range find --session 1 --sheet Sheet1 --find-value "test"
excelcli range replace --session 1 --sheet Sheet1 --find-value "old" --replace-value "new" --replace-all
excelcli range sort --session 1 --sheet Sheet1 --range A1:D10 --sort-columns-json '[{"column":1,"ascending":true}]' --has-headers
```

### Formatting

```bash
excelcli range set-number-format --session 1 --sheet Sheet1 --range B2:B10 --format-code "$#,##0.00"
excelcli range set-font --session 1 --sheet Sheet1 --range A1 --bold --font-size 14 --font-color "#0000FF"
excelcli range set-fill --session 1 --sheet Sheet1 --range A1:D1 --fill-color "#FFFF00"
excelcli range set-borders --session 1 --sheet Sheet1 --range A1:D10 --border-style thin --border-color "#000000"
excelcli range set-alignment --session 1 --sheet Sheet1 --range A1 --horizontal-alignment center --wrap-text
```

### Data Validation

```bash
excelcli range set-validation --session 1 --sheet Sheet1 --range B2:B100 --validation-type list --validation-formula1 "Yes,No,Maybe"
excelcli range set-validation --session 1 --sheet Sheet1 --range C2:C100 --validation-type whole --validation-operator between --validation-formula1 1 --validation-formula2 100
excelcli range clear-validation --session 1 --sheet Sheet1 --range B2:B100
```

### Hyperlinks

```bash
excelcli range set-hyperlink --session 1 --sheet Sheet1 --cell A1 --url "https://example.com" --display-text "Click here"
excelcli range get-hyperlinks --session 1 --sheet Sheet1 --range A1:A10
```

### Worksheets

```bash
excelcli sheet list --session 1
excelcli sheet create --session 1 --name "NewSheet"
excelcli sheet rename --session 1 --sheet Sheet1 --name "DataSheet"
excelcli sheet copy --session 1 --sheet Sheet1 --name "Sheet1_Copy"
excelcli sheet delete --session 1 --sheet OldSheet
excelcli sheet set-tab-color --session 1 --sheet Sheet1 --color "#FF0000"
excelcli sheet set-visibility --session 1 --sheet Hidden --visibility hidden
```

### Tables

```bash
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
```

### Named Ranges

```bash
excelcli namedrange list --session 1
excelcli namedrange create --session 1 --name "TaxRate" --refers-to "=0.25"
excelcli namedrange create --session 1 --name "DataRange" --refers-to "=Sheet1!$A$1:$D$100"
excelcli namedrange update --session 1 --name "TaxRate" --refers-to "=0.30"
excelcli namedrange delete --session 1 --name "OldName"
```

### Power Query

```bash
excelcli powerquery list --session 1
excelcli powerquery get --session 1 --name "SalesQuery"
excelcli powerquery create --session 1 --name "CsvImport" --m-code 'let Source = Csv.Document(File.Contents("C:\Data\sales.csv")) in Source' --load-destination worksheet
excelcli powerquery update --session 1 --name "CsvImport" --m-code 'let Source = Csv.Document(File.Contents("C:\Data\sales_new.csv")) in Source'
excelcli powerquery refresh --session 1 --name "CsvImport" --timeout 120
excelcli powerquery delete --session 1 --name "OldQuery"
```

Load destinations: `worksheet`, `data-model`, `both`, `connection-only`

### Data Model (Power Pivot)

```bash
excelcli datamodel list-tables --session 1
excelcli datamodel list-measures --session 1
excelcli datamodel create-measure --session 1 --table "Sales" --name "TotalRevenue" --dax "SUM(Sales[Amount])"
excelcli datamodel update-measure --session 1 --table "Sales" --name "TotalRevenue" --dax "SUMX(Sales, Sales[Qty] * Sales[Price])"
excelcli datamodel delete-measure --session 1 --table "Sales" --name "OldMeasure"
excelcli datamodel list-relationships --session 1
excelcli datamodel create-relationship --session 1 --from-table "Sales" --from-column "ProductID" --to-table "Products" --to-column "ID"
```

### PivotTables

```bash
excelcli pivottable list --session 1
excelcli pivottable create-from-table --session 1 --source-table "SalesData" --target-sheet "PivotSheet" --target-cell A1 --name "SalesPivot"
excelcli pivottable create-from-datamodel --session 1 --target-sheet "Analysis" --target-cell A1 --name "ModelPivot"
excelcli pivottable add-field --session 1 --name "SalesPivot" --field "Region" --area row
excelcli pivottable add-field --session 1 --name "SalesPivot" --field "Amount" --area data --function sum
excelcli pivottable remove-field --session 1 --name "SalesPivot" --field "Category"
excelcli pivottable refresh --session 1 --name "SalesPivot"
excelcli pivottable delete --session 1 --name "OldPivot"
```

Field areas: `row`, `column`, `data`, `page`
Functions: `sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevP`, `var`, `varP`

### Slicers

```bash
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

```bash
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

```bash
excelcli conditionalformat add-color-scale --session 1 --sheet Sheet1 --range B2:B100 --min-color "#FF0000" --max-color "#00FF00"
excelcli conditionalformat add-data-bar --session 1 --sheet Sheet1 --range C2:C100 --bar-color "#0000FF"
excelcli conditionalformat add-icon-set --session 1 --sheet Sheet1 --range D2:D100 --icon-set "3Arrows"
excelcli conditionalformat add-cell-value --session 1 --sheet Sheet1 --range E2:E100 --operator greaterThan --value 100 --format-fill "#FFFF00"
excelcli conditionalformat clear --session 1 --sheet Sheet1 --range B2:E100
```

### Connections

```bash
excelcli connection list --session 1
excelcli connection get --session 1 --name "SQLServer_Sales"
excelcli connection refresh --session 1 --name "SQLServer_Sales"
excelcli connection refresh-all --session 1
```

### VBA (requires Trust Center enabled)

```bash
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

```bash
# Create workbook
excelcli -q create-empty C:\Reports\Dashboard.xlsx

# Open session
excelcli -q session open C:\Reports\Dashboard.xlsx
# Returns: {"success":true,"sessionId":1,...}

# Import CSV via Power Query
excelcli -q powerquery create --session 1 --name "SalesData" \
  --m-code 'let Source = Csv.Document(File.Contents("C:\Data\sales.csv"),[Delimiter=",",Encoding=65001]) in Source' \
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

```bash
for file in C:\Data\*.xlsx; do
  excelcli -q session open "$file"
  excelcli -q range set-values --session 1 --sheet Summary --range A1 --values-json '[["Updated: 2026-01-28"]]'
  excelcli -q session close --session 1 --save
done
```

### Format Financial Report

```bash
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
