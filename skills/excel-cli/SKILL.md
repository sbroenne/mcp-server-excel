---
name: excel-cli
description: >
  Automate Microsoft Excel on Windows via CLI. Use when creating, reading,
  or modifying Excel workbooks from scripts, CI/CD, or coding agents.
  Supports Power Query, DAX, PivotTables, Tables, Ranges, Charts, VBA.
  Triggers: Excel, spreadsheet, workbook, xlsx, excelcli, CLI automation.
---

# Excel Automation with excelcli

## Preconditions

- Windows host with Microsoft Excel installed (2016+)
- Uses COM interop — does NOT work on macOS or Linux
- Install: `dotnet tool install --global Sbroenne.ExcelMcp.CLI`

## Workflow Checklist

| Step | Command | When |
|------|---------|------|
| 1. Session | `session create/open` | Always first |
| 2. Sheets | `worksheet create/rename` | If needed |
| 3. Write data | See below | If writing values |
| 4. Save & close | `session close --save` | Always last |

> **10+ commands?** Use `excelcli -q batch --input commands.json` — sends all commands in one process with automatic session management. See Rule 8.

**Writing Data (Step 3):**
- `--values` takes a JSON 2D array string: `--values '[["Header1","Header2"],[1,2]]'`
- Write **one row at a time** for reliability: `--range-address A1:B1 --values '[["Name","Age"]]'`
- Strings MUST be double-quoted in JSON: `"text"`. Numbers are bare: `42`
- Always wrap the entire JSON value in single quotes to protect special characters

## CRITICAL RULES (MUST FOLLOW)

> **⚡ Building dashboards or bulk operations?** Skip to **Rule 8: Batch Mode** — it eliminates per-command process overhead and auto-manages session IDs.

### Rule 1: NEVER Ask Clarifying Questions

Execute commands to discover the answer instead:

| DON'T ASK | DO THIS INSTEAD |
|-----------|-----------------|
| "Which file should I use?" | `excelcli -q session list` |
| "What table should I use?" | `excelcli -q table list --session <id>` |
| "Which sheet has the data?" | `excelcli -q worksheet list --session <id>` |

**You have commands to answer your own questions. USE THEM.**

### Rule 2: Always End With a Text Summary

**NEVER end your turn with only a command execution.** After completing all operations, always provide a brief text message confirming what was done. Silent command-only responses are incomplete.

### Rule 3: Session Lifecycle

**Creating vs Opening Files:**
```powershell
# NEW file - use session create
excelcli -q session create C:\path\newfile.xlsx  # Creates file + returns session ID

# EXISTING file - use session open
excelcli -q session open C:\path\existing.xlsx   # Opens file + returns session ID
```

**CRITICAL: Use `session create` for new files. `session open` on non-existent files will fail!**

**CRITICAL: ALWAYS use the session ID returned by `session create` or `session open` in subsequent commands. NEVER guess or hardcode session IDs. The session ID is in the JSON output (e.g., `{"sessionId":"abc123"}`). Parse it and use it.**

```powershell
# Example: capture session ID from output, then use it
excelcli -q session create C:\path\file.xlsx     # Returns JSON with sessionId
excelcli -q range set-values --session <returned-session-id> ...
excelcli -q session close --session <returned-session-id> --save
```

**Unclosed sessions leave Excel processes running, locking files.**

### Rule 4: Data Model Prerequisites

DAX operations require tables in the Data Model:

```powershell
excelcli -q table add-to-data-model --session <id> --table-name Sales  # Step 1
excelcli -q datamodel create-measure --session <id> ...               # Step 2 - NOW works
```

### Rule 5: Power Query Development Lifecycle

**BEST PRACTICE: Test M code before creating permanent queries**

```powershell
# Step 1: Test M code without persisting (catches errors early)
excelcli -q powerquery evaluate --session 1 --m-code-file query.m

# Step 2: Create permanent query with validated code
excelcli -q powerquery create --session 1 --query-name Q1 --m-code-file query.m

# Step 3: Load data to destination
excelcli -q powerquery refresh --session 1 --query-name Q1
```

### Rule 6: Report File Errors Immediately

If you see "File not found" or "Path not found" - STOP and report to user. Don't retry.

### Rule 7: Use Calculation Mode for Bulk Writes

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

### Rule 8: Use Batch Mode for Bulk Operations (10+ commands)

When executing 10+ commands on the same file, use `excelcli batch` to send all commands in a single process launch. This avoids per-process startup overhead and terminal buffer saturation.

```powershell
# Create a JSON file with all commands
@'
[
  {"command": "session.open", "args": {"filePath": "C:\\path\\file.xlsx"}},
  {"command": "range.set-values", "args": {"sheetName": "Sheet1", "rangeAddress": "A1", "values": [["Hello"]]}},
  {"command": "range.set-values", "args": {"sheetName": "Sheet1", "rangeAddress": "A2", "values": [["World"]]}},
  {"command": "session.close", "args": {"save": true}}
]
'@ | Set-Content commands.json

# Execute all commands at once
excelcli -q batch --input commands.json
```

**Key features:**
- **Session auto-capture**: `session.open`/`create` result sessionId auto-injected into subsequent commands — no need to parse and pass session IDs
- **NDJSON output**: One JSON result per line: `{"index": 0, "command": "...", "success": true, "result": {...}}`
- **`--stop-on-error`**: Exit on first failure (default: continue all)
- **`--session <id>`**: Pre-set session ID for all commands (skip session.open)

**Input formats:**
- JSON array from file: `excelcli -q batch --input commands.json`
- NDJSON from stdin: `Get-Content commands.ndjson | excelcli -q batch`

## CLI Command Reference

> Auto-generated from `excelcli --help`. Use these exact parameter names.


### calculationmode

Control Excel recalculation (automatic vs manual). Set manual mode before bulk writes for faster performance, then recalculate once at the end.

**Actions:** `get-mode`, `set-mode`, `calculate`

| Parameter | Description |
|-----------|-------------|
| `--mode` | Target calculation mode (required for: set-mode) |
| `--scope` | Scope: Workbook, Sheet, or Range (required for: calculate) |
| `--sheet-name` | Sheet name (required for Sheet/Range scope) |
| `--range-address` | Range address (required for Range scope) |



### chart

Chart lifecycle - create, read, move, and delete embedded charts. POSITIONING (choose one): - targetRange (PREFERRED): Cell range like 'F2:K15' — positions chart within cells, no point math needed. - left/top: Manual positioning in points (72 points = 1 inch). - Neither: Auto-positions chart below all existing content (used range + other charts). COLLISION DETECTION: All create/move/fit-to-range operations automatically check for overlaps with data and other charts. Warnings are returned in the result message if collisions are detected. Always verify layout with screenshot(capture-sheet) after creating charts. CHART TYPES: 70+ types available including Column, Line, Pie, Bar, Area, XY Scatter. CREATE OPTIONS: - create-from-range: Create from cell range (e.g., 'A1:D10') - create-from-table: Create from Excel Table (uses table's data range) - create-from-pivottable: Create linked PivotChart Use chartconfig for series, titles, legends, styles, placement mode.

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-pivottable`, `delete`, `move`, `fit-to-range`

| Parameter | Description |
|-----------|-------------|
| `--chart-name` | Name of the chart (or shape name) (required for: read, delete, move, fit-to-range) |
| `--sheet-name` | Target worksheet name (required for: create-from-range, create-from-table, create-from-pivottable, fit-to-range) |
| `--source-range-address` | Data range for the chart (e.g., A1:D10) (required for: create-from-range) |
| `--chart-type` | Type of chart to create (required for: create-from-range, create-from-table, create-from-pivottable) |
| `--left` | Left position in points from worksheet edge |
| `--top` | Top position in points from worksheet edge |
| `--width` | Chart width in points |
| `--height` | Chart height in points |
| `--target-range` | Cell range to position chart within (e.g., 'F2:K15'). PREFERRED over left/top. When set, left/top are ignored. |
| `--table-name` | Name of the Excel Table (required for: create-from-table) |
| `--pivot-table-name` | Name of the source PivotTable (required for: create-from-pivottable) |
| `--range-address` | Range to fit the chart to (e.g., A1:D10) (required for: fit-to-range) |



### chartconfig

Chart configuration - data source, series, type, title, axis labels, legend, and styling. SERIES MANAGEMENT: - add-series: Add data series with valuesRange (required) and optional categoryRange - remove-series: Remove series by 1-based index - set-source-range: Replace entire chart data source TITLES AND LABELS: - set-title: Set chart title (empty string hides title) - set-axis-title: Set axis labels (Category, Value, CategorySecondary, ValueSecondary) CHART STYLES: 1-48 (built-in Excel styles with different color schemes) DATA LABELS: Show values, percentages, series/category names. Positions: Center, InsideEnd, InsideBase, OutsideEnd, BestFit. TRENDLINES: Linear, Exponential, Logarithmic, Polynomial (order 2-6), Power, MovingAverage. PLACEMENT MODE: - 1: Move and size with cells - 2: Move but don't size with cells - 3: Don't move or size with cells (free floating) Use chart for lifecycle operations (create, delete, move, fit-to-range).

**Actions:** `set-source-range`, `add-series`, `remove-series`, `set-chart-type`, `set-title`, `set-axis-title`, `get-axis-number-format`, `set-axis-number-format`, `show-legend`, `set-style`, `set-placement`, `set-data-labels`, `get-axis-scale`, `set-axis-scale`, `get-gridlines`, `set-gridlines`, `set-series-format`, `list-trendlines`, `add-trendline`, `delete-trendline`, `set-trendline`

| Parameter | Description |
|-----------|-------------|
| `--chart-name` | Name of the chart (required) |
| `--source-range` | New data source range (e.g., Sheet1!A1:D10) (required for: set-source-range) |
| `--series-name` | Display name for the series (required for: add-series) |
| `--values-range` | Range containing series values (e.g., B2:B10) (required for: add-series) |
| `--category-range` | Optional range for category labels (e.g., A2:A10) |
| `--series-index` | 1-based index of the series to remove (required for: remove-series, set-series-format, list-trendlines, add-trendline, delete-trendline, set-trendline) |
| `--chart-type` | New chart type to apply (required for: set-chart-type) |
| `--title` | Title text to display (required for: set-title, set-axis-title) |
| `--axis` | Which axis to set title for (Category, Value, SeriesAxis) (required for: set-axis-title, get-axis-number-format, set-axis-number-format, get-axis-scale, set-axis-scale, set-gridlines) |
| `--number-format` | Excel number format code (e.g., "$#,##0", "0.00%") (required for: set-axis-number-format) |
| `--visible` | True to show legend, false to hide (required for: show-legend) |
| `--legend-position` | Optional position for the legend |
| `--style-id` | Excel chart style ID (1-48 for most chart types) (required for: set-style) |
| `--placement` | Placement mode: 1=MoveAndSize, 2=Move, 3=FreeFloating (required for: set-placement) |
| `--show-value` | Show data values on labels |
| `--show-percentage` | Show percentage values. Only meaningful for pie and doughnut chart types; setting to true on other chart types has no visual effect. |
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
| `--trendline-type` | Type of trendline (Linear, Exponential, etc.) (required for: add-trendline) |
| `--order` | Polynomial order (2-6, for Polynomial type) |
| `--period` | Moving average period (for MovingAverage type) |
| `--forward` | Periods to extend forward |
| `--backward` | Periods to extend backward |
| `--intercept` | Force trendline through specific Y-intercept |
| `--display-equation` | Display trendline equation on chart |
| `--display-r-squared` | Display R-squared value on chart |
| `--name` | Custom name for the trendline |
| `--trendline-index` | 1-based index of the trendline to delete (required for: delete-trendline, set-trendline) |



### conditionalformat

Conditional formatting - visual rules based on cell values. TYPES: cellValue (requires operatorType+formula1), expression (formula only). Both camelCase and kebab-case accepted. FORMAT: interiorColor/fontColor as #RRGGBB, fontBold/Italic, borderStyle/Color. OPERATORS: equal, notEqual, greater, less, greaterEqual, lessEqual, between, notBetween. For 'between' and 'notBetween', both formula1 and formula2 are required.

**Actions:** `add-rule`, `clear-rules`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Sheet name (empty for active sheet) |
| `--range-address` | Range address (A1 notation or named range) |
| `--rule-type` | Rule type: cellValue (or cell-value), expression, colorScale, dataBar, top10, iconSet, uniqueValues, blanksCondition, timePeriod, aboveAverage. Both camelCase and kebab-case accepted. |
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

Data connections (OLEDB, ODBC, ODC import). TEXT/WEB/CSV: Use powerquery instead. Power Query connections auto-redirect to powerquery. TIMEOUT: 5 min auto-timeout for refresh/load-to.

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

Data Model (Power Pivot) - DAX measures and table management. CRITICAL: WORKSHEET TABLES AND DATA MODEL ARE SEPARATE! - After table append changes, Data Model still has OLD data - MUST call refresh to sync changes - Power Query refresh auto-syncs (no manual refresh needed) PREREQUISITE: Tables must be added to the Data Model first. Use table add-to-datamodel for worksheet tables, or powerquery to import and load data directly to the Data Model. DAX MEASURES: - Create with DAX formulas like 'SUM(Sales[Amount])' - DAX formulas are auto-formatted on CREATE/UPDATE via Dax.Formatter (SQLBI) - Read operations return raw DAX as stored DAX EVALUATE QUERIES: - Use evaluate to execute DAX EVALUATE queries against the Data Model - Returns tabular results from queries like 'EVALUATE TableName' - Supports complex DAX: SUMMARIZE, FILTER, CALCULATETABLE, TOPN, etc. DMV (DYNAMIC MANAGEMENT VIEW) QUERIES: - Use execute-dmv to query Data Model metadata via SQL-like syntax - Syntax: SELECT * FROM $SYSTEM.SchemaRowset (ONLY SELECT * supported) - Use DISCOVER_SCHEMA_ROWSETS to list all available DMVs Use datamodelrel for relationships between tables.

**Actions:** `list-tables`, `list-columns`, `read-table`, `read-info`, `list-measures`, `read`, `delete-measure`, `delete-table`, `rename-table`, `refresh`, `create-measure`, `update-measure`, `evaluate`, `execute-dmv`

| Parameter | Description |
|-----------|-------------|
| `--table-name` | Name of the table to list columns from (required for: list-columns, read-table, delete-table, create-measure) |
| `--measure-name` | Name of the measure to get (required for: read, delete-measure, create-measure, update-measure) |
| `--old-name` | Current name of the table (required for: rename-table) |
| `--new-name` | New name for the table (required for: rename-table) |
| `--timeout` | Optional: Timeout for the refresh operation |
| `--dax-formula` | DAX formula for the measure (will be auto-formatted) (required for: create-measure) |
| `--format-type` | Optional: Format type (Currency, Decimal, Percentage, General) |
| `--description` | Optional: Description of the measure |
| `--dax-query` | DAX EVALUATE query (e.g., "EVALUATE 'TableName'" or "EVALUATE SUMMARIZE(...)") (required for: evaluate) |
| `--dmv-query` | DMV query in SQL-like syntax (e.g., "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES") (required for: execute-dmv) |



### datamodelrelationship

Data Model relationships - link tables for cross-table DAX calculations. CRITICAL: Deleting or recreating tables removes ALL their relationships. Use list-relationships before table operations to backup, then recreate relationships after schema changes. RELATIONSHIP REQUIREMENTS: - Both tables must exist in the Data Model first - Columns must have compatible data types - fromTable/fromColumn = many-side (detail table, foreign key) - toTable/toColumn = one-side (lookup table, primary key) ACTIVE VS INACTIVE: - Only ONE active relationship can exist between two tables - Use active=false when creating alternative paths - DAX USERELATIONSHIP() activates inactive relationships

**Actions:** `list-relationships`, `read-relationship`, `create-relationship`, `update-relationship`, `delete-relationship`

| Parameter | Description |
|-----------|-------------|
| `--from-table` | Source table name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--from-column` | Source column name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--to-table` | Target table name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--to-column` | Target column name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--active` | Whether the relationship should be active (default: true) (required for: update-relationship) |



### diag

Diagnostic commands for testing CLI/MCP infrastructure without Excel. These commands validate parameter parsing, routing, JSON serialization, and error handling — no Excel COM session needed.

**Actions:** `ping`, `echo`, `validate-params`

| Parameter | Description |
|-----------|-------------|
| `--message` | The message to echo back (required) (required for: echo) |
| `--tag` | Optional tag to include in the response |
| `--name` | Required name parameter (required for: validate-params) |
| `--count` | Required integer parameter (required for: validate-params) |
| `--label` | Optional label parameter |
| `--verbose` | Optional boolean flag (default: false) |



### namedrange

Named ranges for formulas/parameters. CREATE/UPDATE: value is cell reference (e.g., 'Sheet1!$A$1'). WRITE: value is data to store. TIP: range(rangeAddress=namedRangeName) for bulk data read/write.

**Actions:** `list`, `write`, `read`, `update`, `create`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--name` | Name of the named range (required for: write, read, update, create, delete) |
| `--value` | Value to set (required for: write) |
| `--reference` | New cell reference (e.g., Sheet1!$A$1:$B$10) (required for: update, create) |



### pivottable

PivotTable lifecycle management: create from various sources, list, read details, refresh, and delete. Use pivottablefield for field operations, pivottablecalc for calculated fields and layout. BEST PRACTICE: Use 'list' before creating. Prefer 'refresh' or field modifications over delete+recreate. Delete+recreate loses field configurations, filters, sorting, and custom layouts. REFRESH: Call 'refresh' after configuring fields with pivottablefield to update the visual display. This is especially important for OLAP/Data Model PivotTables where field operations are structural only and don't automatically trigger a visual refresh. CREATE OPTIONS: - 'create-from-range': Use source sheet and range address for data range - 'create-from-table': Use an Excel Table (ListObject) as source - 'create-from-datamodel': Use a Power Pivot Data Model table as source

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-datamodel`, `delete`, `refresh`

| Parameter | Description |
|-----------|-------------|
| `--pivot-table-name` | Name of the PivotTable (required for: read, create-from-range, create-from-table, create-from-datamodel, delete, refresh) |
| `--source-sheet` | Source worksheet name (required for: create-from-range) |
| `--source-range` | Source range address (e.g., "A1:F100") (required for: create-from-range) |
| `--destination-sheet` | Destination worksheet name (required for: create-from-range, create-from-table, create-from-datamodel) |
| `--destination-cell` | Destination cell address (e.g., "A1") (required for: create-from-range, create-from-table, create-from-datamodel) |
| `--table-name` | Name of the Excel Table (required for: create-from-table, create-from-datamodel) |
| `--timeout` | Optional timeout for the refresh operation |



### pivottablecalc

PivotTable calculated fields/members, layout configuration, and data extraction. Use pivottable for lifecycle, pivottablefield for field placement. CALCULATED FIELDS (for regular PivotTables): - Create custom fields using formulas like '=Revenue-Cost' or '=Quantity*UnitPrice' - Can reference existing fields by name - After creating, use pivottablefield add-value-field to add to Values area - For complex multi-table calculations, prefer DAX measures with datamodel CALCULATED MEMBERS (for OLAP/Data Model PivotTables only): - Create using MDX expressions - Member types: Member, Set, Measure LAYOUT OPTIONS: - 0 = Compact (default, fields in single column) - 1 = Tabular (each field in separate column - best for export/analysis) - 2 = Outline (hierarchical with expand/collapse)

**Actions:** `get-data`, `create-calculated-field`, `list-calculated-fields`, `delete-calculated-field`, `list-calculated-members`, `create-calculated-member`, `delete-calculated-member`, `set-layout`, `set-subtotals`, `set-grand-totals`

| Parameter | Description |
|-----------|-------------|
| `--pivot-table-name` | Name of the PivotTable (required) |
| `--field-name` | Name for the calculated field (required for: create-calculated-field, delete-calculated-field, set-subtotals) |
| `--formula` | Formula using field references (e.g., "=Revenue-Cost") (required for: create-calculated-field, create-calculated-member) |
| `--member-name` | Name for the calculated member (MDX naming format) (required for: create-calculated-member, delete-calculated-member) |
| `--type` | Type of calculated member (Member, Set, or Measure) |
| `--solve-order` | Solve order for calculation precedence (default: 0) |
| `--display-folder` | Display folder path for organizing measures (optional) |
| `--number-format` | Number format code for the calculated member (optional) |
| `--row-layout` | Layout form: 0=Compact, 1=Tabular, 2=Outline (required for: set-layout) |
| `--show-subtotals` | True to show automatic subtotals, false to hide (required for: set-subtotals) |
| `--show-row-grand-totals` | Show row grand totals (bottom summary row) (required for: set-grand-totals) |
| `--show-column-grand-totals` | Show column grand totals (right summary column) (required for: set-grand-totals) |



### pivottablefield

PivotTable field management: add/remove/configure fields, filtering, sorting, and grouping. Use pivottable for lifecycle, pivottablecalc for calculated fields and layout. IMPORTANT: Field operations modify structure only. Call pivottable refresh after configuring fields to update the visual display, especially for OLAP/Data Model PivotTables. FIELD AREAS: - Row fields: Group data by categories (add-row-field) - Column fields: Create column headers (add-column-field) - Value fields: Aggregate numeric data with Sum, Count, Average, etc. (add-value-field) - Filter fields: Add report-level filters (add-filter-field) AGGREGATION FUNCTIONS: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP GROUPING: - Date fields: Group by Days, Months, Quarters, Years (group-by-date) - Numeric fields: Group by ranges with start/end/interval (group-by-numeric) NUMBER FORMAT: Use US format codes like '#,##0.00' for currency or '0.00%' for percentages.

**Actions:** `list-fields`, `add-row-field`, `add-column-field`, `add-value-field`, `add-filter-field`, `remove-field`, `set-field-function`, `set-field-name`, `set-field-format`, `set-field-filter`, `sort-field`, `group-by-date`, `group-by-numeric`

| Parameter | Description |
|-----------|-------------|
| `--pivot-table-name` | Name of the PivotTable (required) |
| `--field-name` | Name of the field to add (required for: add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, set-field-function, set-field-name, set-field-format, set-field-filter, sort-field, group-by-date, group-by-numeric) |
| `--position` | Optional position in row area (1-based) |
| `--aggregation-function` | Aggregation function (for Regular and OLAP auto-create mode) (required for: set-field-function) |
| `--custom-name` | Optional custom name for the field/measure (required for: set-field-name) |
| `--number-format` | Number format string (required for: set-field-format) |
| `--selected-values` | Values to show (others will be hidden) (required for: set-field-filter) |
| `--direction` | Sort direction |
| `--interval` | Grouping interval (Months, Quarters, Years) (required for: group-by-date) |
| `--start` | Starting value (null = use field minimum) |
| `--end-value` | Ending value (null = use field maximum) |
| `--interval-size` | Size of each group (e.g., 100 for groups of 100) (required for: group-by-numeric) |



### powerquery

Power Query M code and data loading. TEST-FIRST DEVELOPMENT WORKFLOW (BEST PRACTICE): 1. evaluate - Test M code WITHOUT persisting (catches syntax errors, validates sources, shows data preview) 2. create/update - Store VALIDATED query in workbook 3. refresh/load-to - Load data to destination Skip evaluate only for trivial literal tables. IF CREATE/UPDATE FAILS: Use evaluate to get the actual M engine error message, fix code, retry. DATETIME COLUMNS: Always include Table.TransformColumnTypes() in M code to set column types explicitly. Without explicit types, dates may be stored as numbers and Data Model relationships may fail. DESTINATIONS: 'worksheet' (default), 'data-model' (for DAX), 'both', 'connection-only'. Use 'data-model' to load to Power Pivot, then use datamodel to create DAX measures. TARGET CELL: targetCellAddress places tables without clearing sheet. TIMEOUT: 5 min auto-timeout for refresh/load. For network queries, use timeout=120 or higher. timeout=0 is INVALID - must be greater than zero.

**Actions:** `list`, `view`, `refresh`, `get-load-config`, `delete`, `create`, `update`, `load-to`, `refresh-all`, `rename`, `unload`, `evaluate`

| Parameter | Description |
|-----------|-------------|
| `--query-name` | Name of the query to view (required for: view, refresh, get-load-config, delete, create, update, load-to, unload) |
| `--timeout` | Maximum time to wait for refresh (required for: refresh) |
| `--m-code` | Raw M code (inline string) (required for: create, update, evaluate) |
| `--load-destination` | Load destination mode |
| `--target-sheet` | Target worksheet name (required for LoadToTable and LoadToBoth; defaults to query name when omitted) |
| `--target-cell-address` | Optional target cell address for worksheet loads (e.g., "B5"). Required when loading to an existing worksheet with other data. |
| `--refresh` | Whether to refresh data after update (default: true) |
| `--old-name` | Current name of the query (required for: rename) |
| `--new-name` | New name for the query (required for: rename) |



### range

Core range operations: get/set values and formulas, copy ranges, clear content, and discover data regions. Use rangeedit for insert/delete/find/sort. Use rangeformat for styling/validation. Use rangelink for hyperlinks and cell protection. Calculation mode and explicit recalculation are handled by calculationmode. BEST PRACTICE: Use 'get-values' to check existing data before overwriting. Use 'clear-contents' (not 'clear-all') to preserve cell formatting when clearing data. set-values preserves existing formatting; use set-number-format after if format change needed. DATA FORMAT: values and formulas are 2D JSON arrays representing rows and columns. Example: [[row1col1, row1col2], [row2col1, row2col2]] Single cell returns [[value]] (always 2D). REQUIRED PARAMETERS: - sheetName + rangeAddress for cell operations (e.g., sheetName='Sheet1', rangeAddress='A1:D10') - For named ranges, use sheetName='' (empty string) and rangeAddress='MyNamedRange' COPY OPERATIONS: Specify source and target sheet/range for copy operations. NUMBER FORMATS: Use US locale format codes (e.g., '#,##0.00', 'mm/dd/yyyy', '0.00%').

**Actions:** `get-values`, `set-values`, `get-formulas`, `set-formulas`, `clear-all`, `clear-contents`, `clear-formats`, `copy`, `copy-values`, `copy-formulas`, `get-number-formats`, `set-number-format`, `set-number-formats`, `get-used-range`, `get-current-region`, `get-info`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Name of the worksheet containing the range - REQUIRED for cell addresses, use empty string for named ranges only (required for: get-values, set-values, get-formulas, set-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-used-range, get-current-region, get-info) |
| `--range-address` | Cell range address (e.g., 'A1', 'A1:D10', 'B:D') or named range name (e.g., 'SalesData') (required for: get-values, set-values, get-formulas, set-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-info) |
| `--values` | 2D array of values to set - rows are outer array, columns are inner array (e.g., [[1,2,3],[4,5,6]] for 2 rows x 3 cols). Optional if valuesFile is provided. |
| `--values-file` | Path to a JSON or CSV file containing the values. JSON: 2D array. CSV: rows/columns. Alternative to inline values parameter. |
| `--formulas` | 2D array of formulas to set - include '=' prefix (e.g., [['=A1+B1', '=SUM(A:A)'], ['=C1*2', '=AVERAGE(B:B)']]). Optional if formulasFile is provided. |
| `--formulas-file` | Path to a JSON file containing the formulas as a 2D array. Alternative to inline formulas parameter. |
| `--source-sheet` | Source worksheet name for copy operations (required for: copy, copy-values, copy-formulas) |
| `--source-range` | Source range address for copy operations (e.g., 'A1:D10') (required for: copy, copy-values, copy-formulas) |
| `--target-sheet` | Target worksheet name for copy operations (required for: copy, copy-values, copy-formulas) |
| `--target-range` | Target range address - can be single cell for paste destination (e.g., 'A1') (required for: copy, copy-values, copy-formulas) |
| `--format-code` | Number format code in US locale (e.g., '#,##0.00' for numbers, 'mm/dd/yyyy' for dates, '0.00%' for percentages, 'General' for default, '@' for text) (required for: set-number-format) |
| `--formats` | 2D array of format codes - same dimensions as target range (e.g., [['#,##0.00', '0.00%'], ['mm/dd/yyyy', 'General']]). Optional if formatsFile is provided. |
| `--formats-file` | Path to a JSON file containing 2D array of format codes. Alternative to inline formats parameter. |
| `--cell-address` | Single cell address (e.g., 'B5') - expands to contiguous data region around this cell (required for: get-current-region) |



### rangeedit

Range editing operations: insert/delete cells, rows, and columns; find/replace text; sort data. Use range for values/formulas/copy/clear operations. INSERT/DELETE CELLS: Specify shift direction to control how surrounding cells move. - Insert: 'Down' or 'Right' - Delete: 'Up' or 'Left' INSERT/DELETE ROWS: Use row range like '5:10' to insert/delete rows 5-10. INSERT/DELETE COLUMNS: Use column range like 'B:D' to insert/delete columns B-D. FIND/REPLACE: Search within the specified range with optional case/cell matching. - Find returns up to 10 matching cell addresses with total count. - Replace modifies all matches by default. SORT: Specify sortColumns as array of {columnIndex: 1, ascending: true} objects. Column indices are 1-based relative to the range.

**Actions:** `insert-cells`, `delete-cells`, `insert-rows`, `delete-rows`, `insert-columns`, `delete-columns`, `find`, `replace`, `sort`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Name of the worksheet containing the range (required) |
| `--range-address` | Cell range address where cells will be inserted (e.g., 'A1:D10') (required) |
| `--insert-shift` | Direction to shift existing cells: 'Down' or 'Right' (required for: insert-cells) |
| `--delete-shift` | Direction to shift remaining cells: 'Up' or 'Left' (required for: delete-cells) |
| `--search-value` | Text or value to search for (required for: find) |
| `--find-options` | Search options: matchCase (default: false), matchEntireCell (default: false), searchFormulas (default: true) (required for: find) |
| `--find-value` | Text or value to search for (required for: replace) |
| `--replace-value` | Text or value to replace matches with (required for: replace) |
| `--replace-options` | Replace options: matchCase (default: false), matchEntireCell (default: false), replaceAll (default: true) (required for: replace) |
| `--sort-columns` | Array of sort specifications: [{columnIndex: 1, ascending: true}, ...] - columnIndex is 1-based relative to range (required for: sort) |
| `--has-headers` | Whether the range has a header row to exclude from sorting (default: true) |



### rangeformat

Range formatting operations: apply styles, set fonts/colors/borders, add data validation, merge cells, auto-fit dimensions. Use range tool for values/formulas/copy/clear operations. set-style: Apply a named Excel style (Heading 1, Good, Bad, Neutral, Normal). Best for semantic status labels (Good/Bad/Neutral have fill colours and are theme-aware) and document hierarchy (Heading 1/2/3). NOTE: Heading styles do NOT apply a fill colour — use format-range when you need a coloured header row. format-range: Apply any combination of bold, fillColor, fontColor, alignment, borders. Required whenever you need a fill colour or custom branding. Pass ALL desired properties in a SINGLE call — do not call format-range multiple times for the same range. COLORS: Hex '#RRGGBB' (e.g., '#FF0000' for red, '#00FF00' for green) FONT: size in points (e.g., 12, 14, 16), alignment: 'left', 'center', 'right' / 'top', 'middle', 'bottom' DATA VALIDATION: Restrict cell input with validation rules: - Types: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom' - For list validation, formula1 is the list source (e.g., '=$A$1:$A$10' or '"Option1,Option2,Option3"') - Operators: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual' MERGE: Combines cells into one. Only top-left cell value is preserved.

**Actions:** `set-style`, `get-style`, `format-range`, `validate-range`, `get-validation`, `remove-validation`, `auto-fit-columns`, `auto-fit-rows`, `merge-cells`, `unmerge-cells`, `get-merge-info`, `set-column-width`, `set-row-height`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Name of the worksheet containing the range (required) |
| `--range-address` | Cell range address (e.g., 'A1:D10') (required) |
| `--style-name` | Built-in or custom style name (e.g., 'Heading 1', 'Good', 'Bad', 'Currency', 'Percent'). Use 'Normal' to reset. (required for: set-style) |
| `--font-name` | Font family name (e.g., 'Arial', 'Calibri', 'Times New Roman') |
| `--font-size` | Font size in points (e.g., 10, 11, 12, 14, 16) |
| `--bold` | Whether to apply bold formatting |
| `--italic` | Whether to apply italic formatting |
| `--underline` | Whether to apply underline formatting |
| `--font-color` | Font (foreground) color as hex '#RRGGBB' (e.g., '#FF0000' for red) |
| `--fill-color` | Cell fill (background) color as hex '#RRGGBB' (e.g., '#FFFF00' for yellow) |
| `--border-style` | Border line style: 'continuous', 'dash', 'dot', 'dashdot', 'dashdotdot', 'double', 'slantdashdot', 'none' |
| `--border-color` | Border color as hex '#RRGGBB' |
| `--border-weight` | Border weight: 'hairline', 'thin', 'medium', 'thick' |
| `--horizontal-alignment` | Horizontal text alignment: 'left', 'center', 'right', 'justify', 'fill' |
| `--vertical-alignment` | Vertical text alignment: 'top', 'center' (or 'middle'), 'bottom', 'justify' |
| `--wrap-text` | Whether to wrap text within cells |
| `--orientation` | Text rotation in degrees (-90 to 90, or 255 for vertical) |
| `--validation-type` | Data validation type: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom' (required for: validate-range) |
| `--validation-operator` | Validation comparison operator: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual' |
| `--formula1` | First validation formula/value - for list validation use range '=$A$1:$A$10' or inline '"A,B,C"' |
| `--formula2` | Second validation formula/value - required only for 'between' and 'notBetween' operators |
| `--show-input-message` | Whether to show input message when cell is selected (default: false) |
| `--input-title` | Title for the input message popup |
| `--input-message` | Text for the input message popup |
| `--show-error-alert` | Whether to show error alert on invalid input (default: true) |
| `--error-style` | Error alert style: 'stop' (prevents entry), 'warning' (allows override), 'information' (allows entry) |
| `--error-title` | Title for the error alert popup |
| `--error-message` | Text for the error alert popup |
| `--ignore-blank` | Whether to allow blank cells in validation (default: true) |
| `--show-dropdown` | Whether to show dropdown arrow for list validation (default: true) |
| `--column-width` | Width in points (1 point = 1/72 inch, approx 0.35mm). Standard width ~8.43 points. Range: 0.25-409 points. (required for: set-column-width) |
| `--row-height` | Height in points (1 point = 1/72 inch, approx 0.35mm). Default row height ~15 points. Range: 0-409 points. (required for: set-row-height) |



### rangelink

Hyperlink and cell protection operations for Excel ranges. Use range for values/formulas, rangeformat for styling. HYPERLINKS: - 'add-hyperlink': Add a clickable hyperlink to a cell (URL can be web, file, or mailto) - 'remove-hyperlink': Remove hyperlink(s) from cells while keeping the cell content - 'list-hyperlinks': Get all hyperlinks on a worksheet - 'get-hyperlink': Get hyperlink details for a specific cell CELL PROTECTION: - 'set-cell-lock': Lock or unlock cells (only effective when sheet protection is enabled) - 'get-cell-lock': Check if cells are locked Note: Cell locking only takes effect when the worksheet is protected.

**Actions:** `add-hyperlink`, `remove-hyperlink`, `list-hyperlinks`, `get-hyperlink`, `set-cell-lock`, `get-cell-lock`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Name of the worksheet (required) |
| `--cell-address` | Single cell address (e.g., 'A1') (required for: add-hyperlink, get-hyperlink) |
| `--url` | Hyperlink URL (web: 'https://...', file: 'file:///...', email: 'mailto:...') (required for: add-hyperlink) |
| `--display-text` | Text to display in the cell (optional, defaults to URL) |
| `--tooltip` | Tooltip text shown on hover (optional) |
| `--range-address` | Cell range address to remove hyperlinks from (e.g., 'A1:D10') (required for: remove-hyperlink, set-cell-lock, get-cell-lock) |
| `--locked` | Lock status: true = locked (protected when sheet protection enabled), false = unlocked (editable) (required for: set-cell-lock) |



### screenshot

Capture Excel worksheet content as images for visual verification. Uses Excel's built-in rendering (CopyPicture) to capture ranges as PNG images. Captures formatting, conditional formatting, charts, and all visual elements. ACTIONS: - capture: Capture a specific range as an image - capture-sheet: Capture the entire used area of a worksheet RETURNS: Base64-encoded image data with dimensions metadata. For MCP: returned as native ImageContent (no file handling needed). For CLI: use --output <path> to save the image directly to a PNG/JPEG file instead of returning base64 inline. Quality defaults to Medium (JPEG 75% scale) which is 4-8x smaller than High (PNG). Use High only when fine detail inspection is needed.

**Actions:** `capture`, `capture-sheet`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Worksheet name (null for active sheet) |
| `--range-address` | Range to capture (e.g., "A1:F20") |
| `--quality` | Image quality: Medium (default, JPEG 75% scale), High (PNG full scale), Low (JPEG 50% scale) |



### sheet

Worksheet lifecycle management: create, rename, copy, delete, move, list sheets. Use range for data operations. Use sheetstyle for tab colors and visibility. ATOMIC OPERATIONS: 'copy-to-file' and 'move-to-file' don't require a session - they open/close files automatically. POSITIONING: For 'move', 'copy-to-file', 'move-to-file' - use 'before' OR 'after' (not both) to position the sheet relative to another. If neither specified, moves to end. NOTE: MCP tool is manually implemented in ExcelWorksheetTool.cs to properly handle mixed session requirements (copy-to-file and move-to-file are atomic and don't need sessions).

**Actions:** `list`, `create`, `rename`, `copy`, `delete`, `move`, `copy-to-file`, `move-to-file`

| Parameter | Description |
|-----------|-------------|
| `--file-path` | Optional file path when batch contains multiple workbooks. If omitted, uses primary workbook. |
| `--sheet-name` | Name for the new worksheet (required for: create, delete, move) |
| `--old-name` | Current name of the worksheet (required for: rename) |
| `--new-name` | New name for the worksheet (required for: rename) |
| `--source-name` | Name of the source worksheet (required for: copy) |
| `--target-name` | Name for the copied worksheet (required for: copy) |
| `--before-sheet` | Optional: Name of sheet to position before |
| `--after-sheet` | Optional: Name of sheet to position after |
| `--source-file` | Full path to the source workbook (required for: copy-to-file, move-to-file) |
| `--source-sheet` | Name of the sheet to copy (required for: copy-to-file, move-to-file) |
| `--target-file` | Full path to the target workbook (required for: copy-to-file, move-to-file) |
| `--target-sheet-name` | Optional: New name for the copied sheet (default: keeps original name) |



### worksheetstyle

Worksheet styling operations for tab colors and visibility. Use sheet for lifecycle operations (create, rename, copy, delete, move). TAB COLORS: Use RGB values (0-255 each) to set custom tab colors for visual organization. VISIBILITY LEVELS: - 'visible': Normal visible sheet - 'hidden': Hidden but accessible via Format > Sheet > Unhide - 'veryhidden': Only accessible via VBA (protection against casual unhiding)

**Actions:** `set-tab-color`, `get-tab-color`, `clear-tab-color`, `set-visibility`, `get-visibility`, `show`, `hide`, `very-hide`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Name of the worksheet to color (required) |
| `--red` | Red color component (0-255) (required for: set-tab-color) |
| `--green` | Green color component (0-255) (required for: set-tab-color) |
| `--blue` | Blue color component (0-255) (required for: set-tab-color) |
| `--visibility` | Visibility level: 'visible', 'hidden', or 'veryhidden' (required for: set-visibility) |



### slicer

Slicer visual filters for PivotTables and Excel Tables. PIVOTTABLE SLICERS: create-slicer, list-slicers, set-slicer-selection, delete-slicer. TABLE SLICERS: create-table-slicer, list-table-slicers, set-table-slicer-selection, delete-table-slicer. NAMING: Auto-generate descriptive names like {FieldName}Slicer (e.g., RegionSlicer). SELECTION: selectedItems as list of strings. Empty list clears filter (shows all items). Set clearFirst=false to add to existing selection.

**Actions:** `create-slicer`, `list-slicers`, `set-slicer-selection`, `delete-slicer`, `create-table-slicer`, `list-table-slicers`, `set-table-slicer-selection`, `delete-table-slicer`

| Parameter | Description |
|-----------|-------------|
| `--pivot-table-name` | Name of the PivotTable to create slicer for (required for: create-slicer) |
| `--field-name` | Name of the field to use for the slicer (required for: create-slicer) |
| `--slicer-name` | Name for the new slicer (required for: create-slicer, set-slicer-selection, delete-slicer, create-table-slicer, set-table-slicer-selection, delete-table-slicer) |
| `--destination-sheet` | Worksheet where slicer will be placed (required for: create-slicer, create-table-slicer) |
| `--position` | Top-left cell position for the slicer (e.g., "H2") (required for: create-slicer, create-table-slicer) |
| `--selected-items` | Items to select (show in PivotTable) (required for: set-slicer-selection, set-table-slicer-selection) |
| `--clear-first` | If true, clears existing selection before setting new items (default: true) |
| `--table-name` | Name of the Excel Table (required for: create-table-slicer) |
| `--column-name` | Name of the column to use for the slicer (required for: create-table-slicer) |



### table

Excel Tables (ListObjects) - lifecycle and data operations. Tables provide structured references, automatic formatting, and Data Model integration. BEST PRACTICE: Use 'list' to check existing tables before creating. Prefer 'append'/'resize'/'rename' over delete+recreate to preserve references. WARNING: Deleting tables used as PivotTable sources or in Data Model relationships will break those objects. DATA MODEL WORKFLOW: To analyze worksheet data with DAX/Power Pivot: 1. Create or identify an Excel Table on a worksheet 2. Use 'add-to-datamodel' to add the table to Power Pivot 3. Then use datamodel to create DAX measures on it DAX-BACKED TABLES: Create tables populated by DAX EVALUATE queries: - 'create-from-dax': Create a new table backed by a DAX query (e.g., SUMMARIZE, FILTER) - 'update-dax': Update the DAX query for an existing DAX-backed table - 'get-dax': Get the DAX query info for a table (check if it's DAX-backed) Related: tablecolumn (filter/sort/columns), datamodel (DAX measures, evaluate queries)

**Actions:** `list`, `create`, `rename`, `delete`, `read`, `resize`, `toggle-totals`, `set-column-total`, `append`, `get-data`, `set-style`, `add-to-data-model`, `create-from-dax`, `update-dax`, `get-dax`

| Parameter | Description |
|-----------|-------------|
| `--sheet-name` | Name of the worksheet to create the table on (required for: create, create-from-dax) |
| `--table-name` | Name for the new table (must be unique in workbook) (required for: create, rename, delete, read, resize, toggle-totals, set-column-total, append, get-data, set-style, add-to-data-model, create-from-dax, update-dax, get-dax) |
| `--range-address` | Cell range address for the table (e.g., 'A1:D10') (required for: create) |
| `--has-headers` | True if first row contains column headers (default: true) |
| `--table-style` | Table style name (e.g., 'TableStyleMedium2', 'TableStyleLight1'). Optional. (required for: set-style) |
| `--new-name` | New name for the table (must be unique in workbook) (required for: rename) |
| `--new-range` | New range address (e.g., 'A1:F20') (required for: resize) |
| `--show-totals` | True to show totals row, false to hide (required for: toggle-totals) |
| `--column-name` | Name of the column to set total function on (required for: set-column-total) |
| `--total-function` | Totals function name: Sum, Count, Average, Min, Max, CountNums, StdDev, Var, None (required for: set-column-total) |
| `--rows` | 2D array of row data to append - column order must match table columns. Optional if rowsFile is provided. |
| `--rows-file` | Path to a JSON or CSV file containing the rows to append. JSON: 2D array. CSV: rows/columns. Alternative to inline rows parameter. |
| `--visible-only` | True to return only visible (non-filtered) rows; false for all rows (default: false) |
| `--strip-bracket-column-names` | When true, renames source table columns that contain literal bracket characters (removes brackets) before adding to the Data Model. This modifies the Excel table column headers in the worksheet. |
| `--dax-query` | DAX EVALUATE query (e.g., 'EVALUATE Sales' or 'EVALUATE SUMMARIZE(...)') (required for: create-from-dax, update-dax) |
| `--target-cell` | Target cell address for table placement (default: 'A1') |



### tablecolumn

Table column, filtering, and sorting operations for Excel Tables (ListObjects). Use table for table-level lifecycle and data operations. FILTERING: - 'apply-filter': Simple criteria filter (e.g., ">100", "=Active", "<>Closed") - 'apply-filter-values': Filter by exact values (provide list of values to include) - 'clear-filters': Remove all active filters - 'get-filters': See current filter state SORTING: - 'sort': Single column sort (ascending/descending) - 'sort-multi': Multi-column sort (provide list of {columnName, ascending} objects) COLUMN MANAGEMENT: - 'add-column'/'remove-column'/'rename-column': Modify table structure NUMBER FORMATS: Use US locale format codes (e.g., '#,##0.00', '0%', 'yyyy-mm-dd')

**Actions:** `apply-filter`, `apply-filter-values`, `clear-filters`, `get-filters`, `add-column`, `remove-column`, `rename-column`, `get-structured-reference`, `sort`, `sort-multi`, `get-column-number-format`, `set-column-number-format`

| Parameter | Description |
|-----------|-------------|
| `--table-name` | Name of the Excel table (required) |
| `--column-name` | Name of the column to filter (required for: apply-filter, apply-filter-values, add-column, remove-column, sort, get-column-number-format, set-column-number-format) |
| `--criteria` | Filter criteria string (e.g., '>100', '=Active', '<>Closed') (required for: apply-filter) |
| `--values` | List of exact values to include in the filter (required for: apply-filter-values) |
| `--position` | 1-based column position (optional, defaults to end of table) |
| `--old-name` | Current column name (required for: rename-column) |
| `--new-name` | New column name (required for: rename-column) |
| `--region` | Table region: 'Data', 'Headers', 'Totals', or 'All' (required for: get-structured-reference) |
| `--ascending` | Sort order: true = ascending (A-Z, 0-9), false = descending (default: true) |
| `--sort-columns` | List of sort specifications: [{columnName: 'Col1', ascending: true}, ...] - applied in order (required for: sort-multi) |
| `--format-code` | Number format code in US locale (e.g., '#,##0.00', '0%', 'yyyy-mm-dd') (required for: set-column-number-format) |



### vba

VBA scripts (requires .xlsm and VBA trust enabled). PREREQUISITES: - Workbook must be macro-enabled (.xlsm) - VBA trust must be enabled for automation RUN: procedureName format is 'Module.Procedure' (e.g., 'Module1.MySub').

**Actions:** `list`, `view`, `import`, `update`, `run`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--module-name` | Name of the VBA module (required for: view, import, update, delete) |
| `--vba-code` | VBA code to import (required for: import, update) |
| `--procedure-name` | Name of the procedure to run (e.g., "Module1.MySub") (required for: run) |
| `--timeout` | Optional timeout for execution |
| `--parameters` | Optional parameters to pass to the procedure (required for: run) |



### window

Control Excel window visibility, position, state, and status bar. Use to show/hide Excel, bring it to front, reposition, or maximize/minimize. Set status bar text to give users real-time feedback during operations. VISIBILITY: 'show' makes Excel visible AND brings to front. 'hide' hides Excel. Visibility changes are reflected in session metadata (session list shows updated state). WINDOW STATE values: 'normal', 'minimized', 'maximized'. ARRANGE presets: 'left-half', 'right-half', 'top-half', 'bottom-half', 'center', 'full-screen'. STATUS BAR: 'set-status-bar' displays text in Excel's status bar. 'clear-status-bar' restores default.

**Actions:** `show`, `hide`, `bring-to-front`, `get-info`, `set-state`, `set-position`, `arrange`, `set-status-bar`, `clear-status-bar`

| Parameter | Description |
|-----------|-------------|
| `--window-state` | Window state: 'normal', 'minimized', or 'maximized' (required for: set-state) |
| `--left` | Window left position in points |
| `--top` | Window top position in points |
| `--width` | Window width in points |
| `--height` | Window height in points |
| `--preset` | Preset name: 'left-half', 'right-half', 'top-half', 'bottom-half', 'center', 'full-screen' (required for: arrange) |
| `--text` | Status bar text to display (e.g. "Building PivotTable from Sales data...") (required for: set-status-bar) |




## Common Pitfalls

### --values-file Must Be an Existing File

`--values-file` expects a path to an **existing** JSON or CSV file on disk. Do NOT pass inline JSON as the value — the CLI will look for a file at that path and fail with "File not found". If you don't have a file, use `--values` with inline JSON instead.

### --timeout Must Be Greater Than Zero

When using `--timeout`, the value must be a positive integer (seconds). `--timeout 0` is invalid and will error. Omit `--timeout` entirely to use the default (300 seconds for most operations).

### Power Query Operations Are Slow

`powerquery create`, `powerquery refresh`, and `powerquery evaluate` may take 30+ seconds depending on data volume. Either omit `--timeout` (uses 5-minute default) or set a generous value like `--timeout 120`.

### JSON Values Format

`--values` takes a 2D JSON array wrapped in single quotes:
```powershell
# CORRECT: 2D array with single-quote wrapper
--values '[["Name","Age"],["Alice",30],["Bob",25]]'

# WRONG: Not a 2D array
--values '["Alice",30]'

# WRONG: Object instead of array
--values '{"Name":"Alice","Age":30}'
```

### List Parameters Use JSON Arrays

Parameters that accept lists (e.g., `--selected-items` for slicers) require JSON array format:
```powershell
# CORRECT: JSON array with single-quote wrapper
--selected-items '["West","East"]'

# CORRECT: Escaped inner quotes
--selected-items "[\"West\",\"East\"]"

# CORRECT: Clear selection
--selected-items '[]'

# WRONG: Comma-separated string (not valid)
--selected-items "West,East"
```

## Reference Documentation

- [Core execution rules and LLM guidelines](./references/behavioral-rules.md)
- [Common mistakes to avoid](./references/anti-patterns.md)
- [Data Model constraints and patterns](./references/workflows.md)
- [Charts, positioning, and formatting](./references/chart.md)
- [Conditional formatting operations](./references/conditionalformat.md)
- [Dashboard and report best practices](./references/dashboard.md)
- [Data Model/DAX specifics](./references/datamodel.md)
- [DMV query reference for Data Model analysis](./references/dmv-reference.md)
- [Power Query M code syntax reference](./references/m-code-syntax.md)
- [PivotTable operations](./references/pivottable.md)
- [Power Query specifics](./references/powerquery.md)
- [Range operations and number formats](./references/range.md)
- [Screenshot and visual verification](./references/screenshot.md)
- [Slicer operations](./references/slicer.md)
- [Table operations](./references/table.md)
- [Worksheet operations](./references/worksheet.md)
