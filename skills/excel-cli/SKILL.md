---
name: excel-cli
description: >
  Automate Microsoft Excel on Windows via CLI. Use when creating, reading,
  or modifying Excel workbooks from scripts, CI/CD, or coding agents.
  Supports Power Query, DAX, PivotTables, Tables, Ranges, Charts, VBA.
  Triggers: Excel, spreadsheet, workbook, xlsx, excelcli, CLI automation.
compatibility: Windows + Microsoft Excel 2016+ required. Uses COM interop - does NOT work on macOS or Linux.
allowed-tools: Cmd(excelcli:*),PowerShell(excelcli:*)
disable-model-invocation: true
license: MIT
version: 1.0.0
repository: https://github.com/sbroenne/mcp-server-excel
documentation: https://excelmcpserver.dev/
---

# Excel Automation with excelcli

## Workflow Checklist

| Step | Command | When |
|------|---------|------|
| 1. Session | `session create/open` | Always first |
| 2. Sheets | `worksheet create/rename` | If needed |
| 3. Write data | See below | If writing values |
| 4. Save & close | `session close --save` | Always last |

**Writing Data (Step 3):**
- **Simple (1-3 cells):** `--values '[[value]]'` inline works
- **Complex (4+ cells):** Use `--values '[[...]]'` with proper JSON escaping

## CRITICAL RULES (MUST FOLLOW)

### Rule 1: NEVER Ask Clarifying Questions

Execute commands to discover the answer instead:

| DON'T ASK | DO THIS INSTEAD |
|-----------|-----------------|
| "Which file should I use?" | `excelcli -q session list` |
| "What table should I use?" | `excelcli -q table list --session <id>` |
| "Which sheet has the data?" | `excelcli -q worksheet list --session <id>` |

**You have commands to answer your own questions. USE THEM.**

### Rule 2: Session Lifecycle

**Creating vs Opening Files:**
```powershell
# NEW file - use session create
excelcli -q session create C:\path\newfile.xlsx  # Creates file + returns session ID

# EXISTING file - use session open  
excelcli -q session open C:\path\existing.xlsx   # Opens file + returns session ID
```

**CRITICAL: Use `session create` for new files. `session open` on non-existent files will fail!**

```powershell
excelcli -q range set-values --session 1 ...   # Use session ID
excelcli -q session close --session 1 --save   # Save and release
```

**Unclosed sessions leave Excel processes running, locking files.**

### Rule 3: Data Model Prerequisites

DAX operations require tables in the Data Model:

```powershell
excelcli -q table add-to-datamodel --session 1 --table-name Sales  # Step 1
excelcli -q datamodel create-measure --session 1 ...               # Step 2 - NOW works
```

### Rule 4: Power Query Development Lifecycle

**BEST PRACTICE: Test M code before creating permanent queries**

```powershell
# Step 1: Test M code without persisting (catches errors early)
excelcli -q powerquery evaluate --session 1 --mcode-file query.m

# Step 2: Create permanent query with validated code
excelcli -q powerquery create --session 1 --query-name Q1 --mcode-file query.m

# Step 3: Load data to destination
excelcli -q powerquery refresh --session 1 --query-name Q1
```

### Rule 5: Report File Errors Immediately

If you see "File not found" or "Path not found" - STOP and report to user. Don't retry.

### Rule 6: Use Calculation Mode for Bulk Writes

When writing many values/formulas (10+ cells), disable auto-recalc for performance:

```powershell
# 1. Set manual mode
excelcli -q calculationmode set-mode --session 1 --mode manual

# 2. Perform bulk writes (use actual parameter names from --help)
excelcli -q range set-values --session 1 --sheet-name Sheet1 --range-address A1 --values '[["data"]]'

# 3. Recalculate once at end
excelcli -q calculationmode calculate --session 1 --scope workbook

# 4. Restore automatic mode
excelcli -q calculationmode set-mode --session 1 --mode automatic
```

## Quick Reference

| Task | Command |
|------|---------|
| **Create new workbook** | `excelcli -q session create <path>` |
| Open existing workbook | `excelcli -q session open <path>` |
| List sheets | `excelcli -q worksheet list --session <id>` |
| **Create sheet** | `excelcli -q worksheet create --session <id> --sheet-name Income` |
| **Rename sheet** | `excelcli -q worksheet rename --session <id> --old-name Sheet1 --new-name Income` |
| Read data | `excelcli -q range get-values --session <id> --sheet-name Sheet1 --range-address A1:D10` |
| Write data | `excelcli -q range set-values --session <id> --sheet-name Sheet1 --range-address A1 --values '[["data"]]'` |
| Create table | `excelcli -q table create --session <id> --sheet-name Sheet1 --range-address A1:D10 --table-name Sales` |
| Add to Data Model | `excelcli -q table add-to-datamodel --session <id> --table-name Sales` |
| Create pivot | `excelcli -q pivottable create-from-table --session <id> --table-name Sales --destination-sheet Analysis` |
| Save & close | `excelcli -q session close --session <id> --save` |

## CLI Command Reference

> Auto-generated from `excelcli --help`. Use these exact parameter names.


### calculationmode

**Actions:** `get-mode`, `set-mode`, `calculate`

| Parameter | Description |
|-----------|-------------|
| `--mode` | Target calculation mode |
| `--scope` | Scope: Workbook, Sheet, or Range |
| `--sheet-name` | Sheet name (required for Sheet/Range scope) |
| `--range-address` | Range address (required for Range scope) |



### chart

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-pivottable`, `delete`, `move`, `fit-to-range`

| Parameter | Description |
|-----------|-------------|
| `--chart-name` | Name of the chart (or shape name) |
| `--sheet-name` |  |
| `--source-range` |  |
| `--chart-type` |  |
| `--left` |  |
| `--top` |  |
| `--width` |  |
| `--height` |  |
| `--table-name` |  |
| `--pivot-table-name` |  |
| `--range-address` |  |



### chartconfig

**Actions:** `set-source-range`, `add-series`, `remove-series`, `set-chart-type`, `set-title`, `set-axis-title`, `get-axis-number-format`, `set-axis-number-format`, `show-legend`, `set-style`, `set-placement`, `set-data-labels`, `get-axis-scale`, `set-axis-scale`, `get-gridlines`, `set-gridlines`, `set-series-format`, `list-trendlines`, `add-trendline`, `delete-trendline`, `set-trendline`

| Parameter | Description |
|-----------|-------------|
| `--chart-name` |  |
| `--source-range` |  |
| `--series-name` |  |
| `--values-range` |  |
| `--category-range` |  |
| `--series-index` |  |
| `--chart-type` |  |
| `--title` |  |
| `--axis` |  |
| `--number-format` |  |
| `--visible` |  |
| `--legend-position` |  |
| `--style-id` |  |
| `--placement` |  |
| `--show-value` |  |
| `--show-percentage` |  |
| `--show-series-name` |  |
| `--show-category-name` |  |
| `--show-bubble-size` |  |
| `--separator` |  |
| `--label-position` |  |
| `--minimum-scale` |  |
| `--maximum-scale` |  |
| `--major-unit` |  |
| `--minor-unit` |  |
| `--show-major` |  |
| `--show-minor` |  |
| `--marker-style` |  |
| `--marker-size` |  |
| `--marker-background-color` |  |
| `--marker-foreground-color` |  |
| `--invert-if-negative` |  |
| `--type` |  |
| `--order` |  |
| `--period` |  |
| `--forward` |  |
| `--backward` |  |
| `--intercept` |  |
| `--display-equation` |  |
| `--display-r-squared` |  |
| `--name` |  |
| `--trendline-index` |  |



### conditionalformat

**Actions:** `add-rule`, `clear-rules`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Sheet name (empty for active sheet) |
| `--range-address` | Range address (A1 notation or named range) |
| `--rule-type` | Rule type: cellValue, expression |
| `--operator-type` | XlFormatConditionOperator: equal, notEqual, greater, less, greaterEqual, lessEqual, between, notBetween |
| `--formula1` | First formula/value for condition |
| `--formula2` | Second formula/value (for between/notBetween) |
| `--interior-color` | Fill color (#RRGGBB or color index) |
| `--interior-pattern` | Interior pattern (1=Solid, -4142=None, 9=Gray50, etc.) |
| `--font-color` | Font color (#RRGGBB or color index) |
| `--font-bold` | Bold font |
| `--font-italic` | Italic font |
| `--border-style` | Border style: none, continuous, dash, dot, etc. |
| `--border-color` | Border color (#RRGGBB or color index) |



### connection

**Actions:** `list`, `view`, `create`, `refresh`, `delete`, `load-to`, `get-properties`, `set-properties`, `test`

| Parameter | Description |
|-----------|-------------|
| `--connection-name` |  |
| `--connection-string` |  |
| `--command-text` |  |
| `--description` |  |
| `--timeout` |  |
| `--sheet-name` |  |
| `--connection-string` |  |
| `--command-text` |  |
| `--background-query` |  |
| `--refresh-on-file-open` |  |
| `--save-password` |  |
| `--refresh-period` |  |



### datamodel

**Actions:** `list-tables`, `list-columns`, `read-table`, `read-info`, `list-measures`, `read`, `delete-measure`, `delete-table`, `rename-table`, `refresh`, `create-measure`, `update-measure`, `evaluate`, `execute-dmv`

| Parameter | Description |
|-----------|-------------|
| `--table-name` | Name of the table to list columns from |
| `--measure-name` | Name of the measure to get |
| `--old-name` | Current name of the table |
| `--new-name` | New name for the table |
| `--timeout` | Optional: Timeout for the refresh operation |
| `--dax-formula` | DAX formula for the measure (will be auto-formatted) |
| `--format-type` | Optional: Format type (Currency, Decimal, Percentage, General) |
| `--description` | Optional: Description of the measure |
| `--dax-query` | DAX EVALUATE query (e.g., "EVALUATE 'TableName'" or "EVALUATE SUMMARIZE(...)") |
| `--dmv-query` | DMV query in SQL-like syntax (e.g., "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES") |



### datamodelrel

**Actions:** `list-relationships`, `read-relationship`, `create-relationship`, `update-relationship`, `delete-relationship`

| Parameter | Description |
|-----------|-------------|
| `--from-table` | Source table name |
| `--from-column` | Source column name |
| `--to-table` | Target table name |
| `--to-column` | Target column name |
| `--active` | Whether the relationship should be active (default: true) |



### namedrange

**Actions:** `list`, `write`, `read`, `update`, `create`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--param-name` |  |
| `--value` |  |
| `--reference` |  |



### pivottable

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-datamodel`, `delete`, `refresh`

| Parameter | Description |
|-----------|-------------|
| `--pivot-table-name` | Name of the PivotTable |
| `--source-sheet` | Source worksheet name |
| `--source-range` | Source range address (e.g., "A1:F100") |
| `--destination-sheet` | Destination worksheet name |
| `--destination-cell` | Destination cell address (e.g., "A1") |
| `--table-name` | Name of the Excel Table |
| `--timeout` | Optional timeout for the refresh operation |



### pivottablecalc

**Actions:** `get-data`, `create-calculated-field`, `list-calculated-fields`, `delete-calculated-field`, `list-calculated-members`, `create-calculated-member`, `delete-calculated-member`, `set-layout`, `set-subtotals`, `set-grand-totals`

| Parameter | Description |
|-----------|-------------|
| `--pivot-table-name` | Name of the PivotTable |
| `--field-name` | Name for the calculated field |
| `--formula` | Formula using field references (e.g., "=Revenue-Cost") |
| `--member-name` | Name for the calculated member (MDX naming format) |
| `--type` | Type of calculated member (Member, Set, or Measure) |
| `--solve-order` | Solve order for calculation precedence (default: 0) |
| `--display-folder` | Display folder path for organizing measures (optional) |
| `--number-format` | Number format code for the calculated member (optional) |
| `--layout-type` | Layout form: 0=Compact, 1=Tabular, 2=Outline |
| `--show-subtotals` | True to show automatic subtotals, false to hide |
| `--show-row-grand-totals` | Show row grand totals (bottom summary row) |
| `--show-column-grand-totals` | Show column grand totals (right summary column) |



### pivottablefield

**Actions:** `list-fields`, `add-row-field`, `add-column-field`, `add-value-field`, `add-filter-field`, `remove-field`, `set-field-function`, `set-field-name`, `set-field-format`, `set-field-filter`, `sort-field`, `group-by-date`, `group-by-numeric`

| Parameter | Description |
|-----------|-------------|
| `--pivot-table-name` | Name of the PivotTable |
| `--field-name` | Name of the field to add |
| `--position` | Optional position in row area (1-based) |
| `--aggregation-function` | Aggregation function (for Regular and OLAP auto-create mode) |
| `--custom-name` | Optional custom name for the field/measure |
| `--number-format` | Number format string |
| `--selected-values` | Values to show (others will be hidden) |
| `--direction` | Sort direction |
| `--interval` | Grouping interval (Months, Quarters, Years) |
| `--start` | Starting value (null = use field minimum) |
| `--end-value` | Ending value (null = use field maximum) |
| `--interval-size` | Size of each group (e.g., 100 for groups of 100) |



### powerquery

**Actions:** `list`, `view`, `refresh`, `get-load-config`, `delete`, `create`, `update`, `load-to`, `refresh-all`, `rename`, `unload`, `evaluate`

| Parameter | Description |
|-----------|-------------|
| `--query-name` |  |
| `--timeout` |  |
| `--m-code` | Raw M code (inline string) |
| `--load-destination` | Load destination mode |
| `--target-sheet` | Target worksheet name (required for LoadToTable and LoadToBoth; defaults to query name when omitted) |
| `--target-cell-address` | Optional target cell address for worksheet loads (e.g., "B5"). Required when loading to an existing worksheet with other data. |
| `--refresh` | Whether to refresh data after update (default: true) |
| `--old-name` | Existing query name |
| `--new-name` | Desired new name |



### range

**Actions:** `get-values`, `set-values`, `get-formulas`, `set-formulas`, `clear-all`, `clear-contents`, `clear-formats`, `copy`, `copy-values`, `copy-formulas`, `get-number-formats`, `set-number-format`, `set-number-formats`, `get-used-range`, `get-current-region`, `get-info`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` |  |
| `--range-address` |  |
| `--values` | 2D array of values to set |
| `--formulas` |  |
| `--source-sheet` |  |
| `--source-range` |  |
| `--target-sheet` |  |
| `--target-range` |  |
| `--format-code` | Excel format code (e.g., "$#,##0.00", "0.00%", "m/d/yyyy", "General", "@") |
| `--formats` |  |
| `--cell-address` |  |



### rangeedit

**Actions:** `insert-cells`, `delete-cells`, `insert-rows`, `delete-rows`, `insert-columns`, `delete-columns`, `find`, `replace`, `sort`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` |  |
| `--range-address` |  |
| `--insert-shift` |  |
| `--delete-shift` |  |
| `--search-value` |  |
| `--find-options` |  |
| `--find-value` |  |
| `--replace-value` |  |
| `--replace-options` |  |
| `--sort-columns` |  |
| `--has-headers` |  |



### rangeformat

**Actions:** `set-style`, `get-style`, `format-range`, `validate-range`, `get-validation`, `remove-validation`, `auto-fit-columns`, `auto-fit-rows`, `merge-cells`, `unmerge-cells`, `get-merge-info`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Sheet name (empty for active sheet) |
| `--range-address` | Range address (e.g., "A1:D10") |
| `--style-name` | Built-in style name (e.g., "Heading 1", "Good", "Currency"). Use "Normal" to reset. |
| `--font-name` |  |
| `--font-size` |  |
| `--bold` |  |
| `--italic` |  |
| `--underline` |  |
| `--font-color` |  |
| `--fill-color` |  |
| `--border-style` |  |
| `--border-color` |  |
| `--border-weight` |  |
| `--horizontal-alignment` |  |
| `--vertical-alignment` |  |
| `--wrap-text` |  |
| `--orientation` |  |
| `--validation-type` |  |
| `--validation-operator` |  |
| `--formula1` |  |
| `--formula2` |  |
| `--show-input-message` |  |
| `--input-title` |  |
| `--input-message` |  |
| `--show-error-alert` |  |
| `--error-style` |  |
| `--error-title` |  |
| `--error-message` |  |
| `--ignore-blank` |  |
| `--show-dropdown` |  |



### rangelink

**Actions:** `add-hyperlink`, `remove-hyperlink`, `list-hyperlinks`, `get-hyperlink`, `set-cell-lock`, `get-cell-lock`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` |  |
| `--cell-address` |  |
| `--url` |  |
| `--display-text` |  |
| `--tooltip` |  |
| `--range-address` |  |
| `--locked` |  |



### worksheet

**Actions:** `list`, `create`, `rename`, `copy`, `delete`, `move`, `copy-to-file`, `move-to-file`

| Parameter | Description |
|-----------|-------------|
| `--file-path` | Optional file path when batch contains multiple workbooks. If omitted, uses primary workbook. |
| `--sheet-name` | Name for the new worksheet |
| `--old-name` | Current name of the worksheet |
| `--new-name` | New name for the worksheet |
| `--source-name` | Name of the source worksheet |
| `--target-name` | Name for the copied worksheet |
| `--before-sheet` | Optional: Name of sheet to position before |
| `--after-sheet` | Optional: Name of sheet to position after |
| `--source-file` | Full path to the source workbook |
| `--source-sheet` | Name of the sheet to copy |
| `--target-file` | Full path to the target workbook |
| `--target-sheet-name` | Optional: New name for the copied sheet (default: keeps original name) |



### worksheetstyle

**Actions:** `set-tab-color`, `get-tab-color`, `clear-tab-color`, `set-visibility`, `get-visibility`, `show`, `hide`, `very-hide`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Name of the worksheet |
| `--red` | Red component (0-255) |
| `--green` | Green component (0-255) |
| `--blue` | Blue component (0-255) |
| `--visibility` | Visibility level: visible, hidden, or veryhidden |



### slicer

**Actions:** `create-slicer`, `list-slicers`, `set-slicer-selection`, `delete-slicer`, `create-table-slicer`, `list-table-slicers`, `set-table-slicer-selection`, `delete-table-slicer`

| Parameter | Description |
|-----------|-------------|
| `--pivot-table-name` | Name of the PivotTable to create slicer for |
| `--field-name` | Name of the field to use for the slicer |
| `--slicer-name` | Name for the new slicer |
| `--destination-sheet` | Worksheet where slicer will be placed |
| `--position` | Top-left cell position for the slicer (e.g., "H2") |
| `--selected-items` | Items to select (show in PivotTable) |
| `--clear-first` | If true, clears existing selection before setting new items (default: true) |
| `--table-name` | Name of the Excel Table |
| `--column-name` | Name of the column to use for the slicer |



### table

**Actions:** `list`, `create`, `rename`, `delete`, `read`, `resize`, `toggle-totals`, `set-column-total`, `append`, `get-data`, `set-style`, `add-to-data-model`, `create-from-dax`, `update-dax`, `get-dax`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` |  |
| `--table-name` |  |
| `--range` |  |
| `--has-headers` |  |
| `--table-style` |  |
| `--new-name` |  |
| `--new-range` |  |
| `--show-totals` |  |
| `--column-name` |  |
| `--total-function` |  |
| `--rows` |  |
| `--visible-only` | If true, only rows not hidden by filters are returned |
| `--dax-query` | DAX EVALUATE query (e.g., "EVALUATE 'TableName'" or "EVALUATE SUMMARIZE(...)") |
| `--target-cell` | Target cell address for table placement (default: "A1") |



### tablecolumn

**Actions:** `apply-filter`, `apply-filter-values`, `clear-filters`, `get-filters`, `add-column`, `remove-column`, `rename-column`, `get-structured-reference`, `sort`, `sort-multi`, `get-column-number-format`, `set-column-number-format`

| Parameter | Description |
|-----------|-------------|
| `--table-name` |  |
| `--column-name` |  |
| `--criteria` |  |
| `--values` |  |
| `--position` |  |
| `--old-name` |  |
| `--new-name` |  |
| `--region` |  |
| `--ascending` |  |
| `--sort-columns` |  |
| `--format-code` | Excel format code (e.g., "$#,##0.00", "0.00%") |



### vba

**Actions:** `list`, `view`, `import`, `update`, `run`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--module-name` |  |
| `--vba-code` |  |
| `--procedure-name` |  |
| `--timeout` |  |
| `--parameters` |  |




## Reference Documentation

- @references/behavioral-rules.md - Core execution rules and LLM guidelines
- @references/anti-patterns.md - Common mistakes to avoid
- @references/workflows.md - Data Model constraints and patterns

## Installation

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```
