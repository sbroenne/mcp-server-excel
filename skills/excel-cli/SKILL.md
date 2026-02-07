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
- `--values` takes a JSON 2D array string: `--values '[["Header1","Header2"],[1,2]]'`
- Write **one row at a time** for reliability: `--range-address A1:B1 --values '[["Name","Age"]]'`
- Strings MUST be double-quoted in JSON: `"text"`. Numbers are bare: `42`
- Always wrap the entire JSON value in single quotes to protect special characters

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

# 2. Write data row by row for reliability
excelcli -q range set-values --session 1 --sheet-name Sheet1 --range-address A1:B1 --values '[["Name","Amount"]]'
excelcli -q range set-values --session 1 --sheet-name Sheet1 --range-address A2:B2 --values '[["Salary",5000]]'

# 3. Recalculate once at end
excelcli -q calculationmode calculate --session 1 --scope workbook

# 4. Restore automatic mode
excelcli -q calculationmode set-mode --session 1 --mode automatic
```

## CLI Command Reference

> Auto-generated from `excelcli --help`. Use these exact parameter names.


### calculationmode

Control Excel recalculation (automatic vs manual). Set manual mode before bulk writes for faster performance, then recalculate once at the end.

**Actions:** `get-mode`, `set-mode`, `calculate`

| Parameter | Description |
|-----------|-------------|
| `--mode` | Target calculation mode |
| `--scope` | Scope: Workbook, Sheet, or Range |
| `--sheet-name` | Sheet name (required for Sheet/Range scope) |
| `--range-address` | Range address (required for Range scope) |



### chart

Excel chart lifecycle operations - creating, reading, moving, and deleting charts. Supports Regular charts (static, from ranges) and PivotCharts (dynamic, from PivotTables). Use chartconfig command for series, titles, and styling.

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-pivottable`, `delete`, `move`, `fit-to-range`

| Parameter | Description |
|-----------|-------------|
| `--chart-name` | Name of the chart (or shape name) |
| `--sheet-name` | Target worksheet name |
| `--source-range` | Data range for the chart (e.g., A1:D10) |
| `--chart-type` | Type of chart to create |
| `--left` | Left position in points from worksheet edge |
| `--top` | Top position in points from worksheet edge |
| `--width` | Chart width in points |
| `--height` | Chart height in points |
| `--table-name` | Name of the Excel Table |
| `--pivot-table-name` | Name of the source PivotTable |
| `--range-address` | Range to fit the chart to (e.g., A1:D10) |



### chartconfig

Excel chart configuration operations - data sources, titles, axes, styling, trendlines. Use chart command for lifecycle operations (create, delete, move).

**Actions:** `set-source-range`, `add-series`, `remove-series`, `set-chart-type`, `set-title`, `set-axis-title`, `get-axis-number-format`, `set-axis-number-format`, `show-legend`, `set-style`, `set-placement`, `set-data-labels`, `get-axis-scale`, `set-axis-scale`, `get-gridlines`, `set-gridlines`, `set-series-format`, `list-trendlines`, `add-trendline`, `delete-trendline`, `set-trendline`

| Parameter | Description |
|-----------|-------------|
| `--chart-name` | Name of the chart |
| `--source-range` | New data source range (e.g., Sheet1!A1:D10) |
| `--series-name` | Display name for the series |
| `--values-range` | Range containing series values (e.g., B2:B10) |
| `--category-range` | Optional range for category labels (e.g., A2:A10) |
| `--series-index` | 1-based index of the series to remove |
| `--chart-type` | New chart type to apply |
| `--title` | Title text to display |
| `--axis` | Which axis to set title for (Category, Value, SeriesAxis) |
| `--number-format` | Excel number format code (e.g., "$#,##0", "0.00%") |
| `--visible` | True to show legend, false to hide |
| `--legend-position` | Optional position for the legend |
| `--style-id` | Excel chart style ID (1-48 for most chart types) |
| `--placement` | Placement mode: 1=MoveAndSize, 2=Move, 3=FreeFloating |
| `--show-value` | Show data values on labels |
| `--show-percentage` | Show percentage values (pie/doughnut charts) |
| `--show-series-name` | Show series name on labels |
| `--show-category-name` | Show category name on labels |
| `--show-bubble-size` | Show bubble size (bubble charts) |
| `--separator` | Separator string between label components |
| `--label-position` | Position of data labels relative to data points |
| `--minimum-scale` | Minimum axis value (null for auto) |
| `--maximum-scale` | Maximum axis value (null for auto) |
| `--major-unit` | Major gridline interval (null for auto) |
| `--minor-unit` | Minor gridline interval (null for auto) |
| `--show-major` | Show major gridlines (null to keep current) |
| `--show-minor` | Show minor gridlines (null to keep current) |
| `--marker-style` | Marker shape style |
| `--marker-size` | Marker size in points (2-72) |
| `--marker-background-color` | Marker fill color (#RRGGBB) |
| `--marker-foreground-color` | Marker border color (#RRGGBB) |
| `--invert-if-negative` | Invert colors for negative values |
| `--type` | Type of trendline (Linear, Exponential, etc.) |
| `--order` | Polynomial order (2-6, for Polynomial type) |
| `--period` | Moving average period (for MovingAverage type) |
| `--forward` | Periods to extend forward |
| `--backward` | Periods to extend backward |
| `--intercept` | Force trendline through specific Y-intercept |
| `--display-equation` | Display trendline equation on chart |
| `--display-r-squared` | Display R-squared value on chart |
| `--name` | Custom name for the trendline |
| `--trendline-index` | 1-based index of the trendline to delete |



### conditionalformat

Apply conditional formatting rules based on cell values or formulas. Supports color, font, and border formatting with comparison operators.

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

Manage external data connections (ODBC, OLE DB) and their refresh/load behavior. For M code authoring and query logic, use the powerquery command instead.

**Actions:** `list`, `view`, `create`, `refresh`, `delete`, `load-to`, `get-properties`, `set-properties`, `test`

| Parameter | Description |
|-----------|-------------|
| `--connection-name` | Name of the connection to view |
| `--connection-string` | OLEDB or ODBC connection string |
| `--command-text` | SQL query or table name |
| `--description` | Optional description for the connection |
| `--timeout` | Optional timeout for the refresh operation |
| `--sheet-name` | Target worksheet name |
| `--connection-string` | New connection string (null to keep current) |
| `--command-text` | New SQL query or table name (null to keep current) |
| `--background-query` | Run query in background (null to keep current) |
| `--refresh-on-file-open` | Refresh when file opens (null to keep current) |
| `--save-password` | Save password in connection (null to keep current) |
| `--refresh-period` | Auto-refresh interval in minutes (null to keep current) |



### datamodel

Power Pivot Data Model - DAX measures, DAX queries, and DMV introspection. Tables must be added first with table add-to-datamodel. Use datamodelrel for relationships between tables.

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

Define relationships between Data Model tables for cross-table DAX calculations. Relationships link a foreign key column to a primary key column.

**Actions:** `list-relationships`, `read-relationship`, `create-relationship`, `update-relationship`, `delete-relationship`

| Parameter | Description |
|-----------|-------------|
| `--from-table` | Source table name |
| `--from-column` | Source column name |
| `--to-table` | Target table name |
| `--to-column` | Target column name |
| `--active` | Whether the relationship should be active (default: true) |



### namedrange

Named ranges give human-readable aliases to cell ranges or formulas. Use for dynamic references, input parameters, and improving formula readability.

**Actions:** `list`, `write`, `read`, `update`, `create`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--param-name` | Name of the named range |
| `--value` | Value to set |
| `--reference` | New cell reference (e.g., Sheet1!$A$1:$B$10) |



### pivottable

PivotTable lifecycle operations (create, list, read, refresh, delete). Use pivottablefield for field operations, pivottablecalc for calculated fields and layout.

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

PivotTable calculated fields, layout mode (compact/tabular/outline), subtotals, grand totals, and raw data extraction. Use pivottable for lifecycle, pivottablefield for field placement.

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

Place fields into PivotTable areas (rows, columns, values, filters), set aggregation functions, apply formatting, sort, and group by date or numeric intervals.

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

Power Query (M code) management - create, edit, execute, and load queries. Use for ETL operations, data transformation, and connecting to external data sources.

**Actions:** `list`, `view`, `refresh`, `get-load-config`, `delete`, `create`, `update`, `load-to`, `refresh-all`, `rename`, `unload`, `evaluate`

| Parameter | Description |
|-----------|-------------|
| `--query-name` | Name of the query to view |
| `--timeout` | Maximum time to wait for refresh |
| `--m-code` | Raw M code (inline string) |
| `--load-destination` | Load destination mode |
| `--target-sheet` | Target worksheet name (required for LoadToTable and LoadToBoth; defaults to query name when omitted) |
| `--target-cell-address` | Optional target cell address for worksheet loads (e.g., "B5"). Required when loading to an existing worksheet with other data. |
| `--refresh` | Whether to refresh data after update (default: true) |
| `--old-name` | Current name of the query |
| `--new-name` | New name for the query |



### range

Core range data operations - values, formulas, copy, clear, discovery. Single cell is 1x1 range. Named ranges work via rangeAddress parameter. Use rangeedit for insert/delete/find, rangeformat for styling, rangelink for hyperlinks.

**Actions:** `get-values`, `set-values`, `get-formulas`, `set-formulas`, `clear-all`, `clear-contents`, `clear-formats`, `copy`, `copy-values`, `copy-formulas`, `get-number-formats`, `set-number-format`, `set-number-formats`, `get-used-range`, `get-current-region`, `get-info`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Worksheet name (empty for named ranges) |
| `--range-address` | Range address (e.g., A1:C10) or named range name |
| `--values` | 2D array of values to set |
| `--formulas` | 2D array of formula strings |
| `--source-sheet` | Source worksheet name |
| `--source-range` | Source range address |
| `--target-sheet` | Target worksheet name |
| `--target-range` | Target range address (top-left cell) |
| `--format-code` | Excel format code (e.g., "$#,##0.00", "0.00%", "m/d/yyyy", "General", "@") |
| `--formats` | 2D array of format codes |
| `--cell-address` | Cell address to find region around |



### rangeedit

Range edit operations - insert, delete, find, replace, sort rows/columns. Use range command for values/formulas/copy/clear operations.

**Actions:** `insert-cells`, `delete-cells`, `insert-rows`, `delete-rows`, `insert-columns`, `delete-columns`, `find`, `replace`, `sort`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Worksheet name |
| `--range-address` | Range where cells will be inserted |
| `--insert-shift` | Direction to shift existing cells (Down or Right) |
| `--delete-shift` | Direction to shift remaining cells (Up or Left) |
| `--search-value` | Value to find |
| `--find-options` | Search options (case, whole cell, etc.) |
| `--find-value` | Value to find |
| `--replace-value` | Value to replace with |
| `--replace-options` | Replace options (case, whole cell, etc.) |
| `--sort-columns` | Columns to sort by with direction |
| `--has-headers` | True if first row contains headers |



### rangeformat

Range formatting - fonts, colors, borders, number formats, validation, merge, autofit. Use range command for values/formulas/copy/clear operations.

**Actions:** `set-style`, `get-style`, `format-range`, `validate-range`, `get-validation`, `remove-validation`, `auto-fit-columns`, `auto-fit-rows`, `merge-cells`, `unmerge-cells`, `get-merge-info`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Sheet name (empty for active sheet) |
| `--range-address` | Range address (e.g., "A1:D10") |
| `--style-name` | Built-in style name (e.g., "Heading 1", "Good", "Currency"). Use "Normal" to reset. |
| `--font-name` | Font family name (e.g., "Calibri", "Arial") |
| `--font-size` | Font size in points |
| `--bold` | Bold formatting |
| `--italic` | Italic formatting |
| `--underline` | Underline formatting |
| `--font-color` | Font color (#RRGGBB or color name) |
| `--fill-color` | Background fill color (#RRGGBB or color name) |
| `--border-style` | Border line style |
| `--border-color` | Border color (#RRGGBB or color name) |
| `--border-weight` | Border thickness |
| `--horizontal-alignment` | Horizontal text alignment |
| `--vertical-alignment` | Vertical text alignment |
| `--wrap-text` | Wrap text within cell |
| `--orientation` | Text rotation angle (-90 to 90) |
| `--validation-type` | Type of validation (List, WholeNumber, Decimal, Date, etc.) |
| `--validation-operator` | Comparison operator (Between, Equal, GreaterThan, etc.) |
| `--formula1` | First value/formula for validation |
| `--formula2` | Second value for Between/NotBetween operators |
| `--show-input-message` | Show input message when cell selected |
| `--input-title` | Input message title |
| `--input-message` | Input message text |
| `--show-error-alert` | Show error alert on invalid entry |
| `--error-style` | Error alert style (Stop, Warning, Information) |
| `--error-title` | Error alert title |
| `--error-message` | Error alert message |
| `--ignore-blank` | Allow blank entries |
| `--show-dropdown` | Show dropdown for List validation |



### rangelink

Manage hyperlinks (add, remove, list) and cell lock state for worksheet protection. Use range for values/formulas, rangeformat for styling.

**Actions:** `add-hyperlink`, `remove-hyperlink`, `list-hyperlinks`, `get-hyperlink`, `set-cell-lock`, `get-cell-lock`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Worksheet name |
| `--cell-address` | Cell address (e.g., A1) |
| `--url` | URL or file path for the hyperlink |
| `--display-text` | Optional text to display (defaults to URL) |
| `--tooltip` | Optional tooltip on hover |
| `--range-address` | Range address to remove hyperlinks from |
| `--locked` | True to lock cells, false to unlock |



### worksheet

Worksheet lifecycle - create, rename, copy, delete, list, activate worksheets. Use range command for data operations. Use worksheetstyle for tab colors and visibility.

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

Worksheet styling and appearance - tab colors, visibility, freeze panes. Use worksheet command for lifecycle operations (create, rename, copy, delete).

**Actions:** `set-tab-color`, `get-tab-color`, `clear-tab-color`, `set-visibility`, `get-visibility`, `show`, `hide`, `very-hide`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Name of the worksheet |
| `--red` | Red component (0-255) |
| `--green` | Green component (0-255) |
| `--blue` | Blue component (0-255) |
| `--visibility` | Visibility level: visible, hidden, or veryhidden |



### slicer

Slicer visual filters for PivotTables and Excel Tables. Create interactive filter controls that can be shared across multiple PivotTables.

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

Excel Table (ListObject) lifecycle - create, read, resize, rename, delete structured tables. Tables provide structured references, automatic formatting, and Data Model integration.

**Actions:** `list`, `create`, `rename`, `delete`, `read`, `resize`, `toggle-totals`, `set-column-total`, `append`, `get-data`, `set-style`, `add-to-data-model`, `create-from-dax`, `update-dax`, `get-dax`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Worksheet name |
| `--table-name` | Name for the new table |
| `--range` | Range address for the table (e.g., A1:D10) |
| `--has-headers` | True if first row contains headers |
| `--table-style` | Optional table style name (e.g., "TableStyleMedium2") |
| `--new-name` | New table name |
| `--new-range` | New range address (e.g., A1:F20) |
| `--show-totals` | True to show totals row, false to hide |
| `--column-name` | Column name |
| `--total-function` | Function name (Sum, Count, Average, Min, Max, etc.) |
| `--rows` | 2D array of row data to append |
| `--visible-only` | If true, only rows not hidden by filters are returned |
| `--dax-query` | DAX EVALUATE query (e.g., "EVALUATE 'TableName'" or "EVALUATE SUMMARIZE(...)") |
| `--target-cell` | Target cell address for table placement (default: "A1") |



### tablecolumn

Column-level Table operations: AutoFilter, value filters, multi-column sorting, structured references, and number formatting. Use table for table-level operations.

**Actions:** `apply-filter`, `apply-filter-values`, `clear-filters`, `get-filters`, `add-column`, `remove-column`, `rename-column`, `get-structured-reference`, `sort`, `sort-multi`, `get-column-number-format`, `set-column-number-format`

| Parameter | Description |
|-----------|-------------|
| `--table-name` | Name of the Excel table |
| `--column-name` | Name of the column to filter |
| `--criteria` | Filter criteria (e.g., ">100", "=Active") |
| `--values` | List of values to include in filter |
| `--position` | Optional 1-based position (null for last) |
| `--old-name` | Current column name |
| `--new-name` | New column name |
| `--region` | Table region (Data, Headers, Totals, All) |
| `--ascending` | True for ascending, false for descending |
| `--sort-columns` | List of columns with sort direction |
| `--format-code` | Excel format code (e.g., "$#,##0.00", "0.00%") |



### vba

VBA macro management - list, view, create, edit, run VBA modules and procedures. Requires macro-enabled workbooks (.xlsm) and proper trust settings.

**Actions:** `list`, `view`, `import`, `update`, `run`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--module-name` | Name of the VBA module |
| `--vba-code` | VBA code to import |
| `--procedure-name` | Name of the procedure to run (e.g., "Module1.MySub") |
| `--timeout` | Optional timeout for execution |
| `--parameters` | Optional parameters to pass to the procedure |




## Reference Documentation

- @references/behavioral-rules.md - Core execution rules and LLM guidelines
- @references/anti-patterns.md - Common mistakes to avoid
- @references/workflows.md - Data Model constraints and patterns

## Installation

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```
