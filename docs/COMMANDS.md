---

## Workflow Guidance and Next Actions

All commands now display AI-powered workflow guidance after successful operations:

- **Workflow Hint**: Contextual guidance in dim text
- **Suggested Next Actions**: Bulleted list of recommended next steps

**Example Output:**
```
✓ Imported VBA module 'DataProcessor' from processor.vba
Next steps: Review the imported code for any required customizations

Suggested Next Actions:
  • Run the VBA procedure with vba-run
  • Export the module for version control with vba-export
  • Test the module functionality
```

### New Features
- Phase 1 atomic operations: `pq-create`, `pq-update-mcode`, `pq-update-and-refresh`, `pq-unload`, `pq-refresh-all`
- All commands display workflow hints and suggested actions after success

---

# Command Reference

Complete reference for all ExcelMcp.CLI commands.

## File Commands

Create and manage Excel workbooks.

**create-empty** - Create empty Excel workbook

```powershell
excelcli create-empty <file.xlsx|file.xlsm>
```

Essential for coding agents to programmatically create Excel workbooks. Use `.xlsm` extension for macro-enabled workbooks that support VBA.

## Power Query Commands (`pq-*`)

Manage Power Query queries in Excel workbooks.

**pq-list** - List all Power Queries

```powershell
excelcli pq-list <file.xlsx>
```

**pq-view** - View Power Query M code

```powershell
excelcli pq-view <file.xlsx> <query-name>
```

**pq-export** - Export query to file

```powershell
excelcli pq-export <file.xlsx> <query-name> <output.pq>
```

**pq-refresh** - Refresh query data

```powershell
excelcli pq-refresh <file.xlsx> <query-name>
```

Refreshes an existing Power Query to reload data from its source. If the query is connection-only (not loaded anywhere), use the MCP Server `excel_powerquery` tool with the `loadDestination` parameter to apply load configuration during refresh:

```javascript
// MCP Server usage - Apply load destination while refreshing
excel_powerquery(action: "refresh", queryName: "Sales", loadDestination: "worksheet")
excel_powerquery(action: "refresh", queryName: "Sales", loadDestination: "data-model")
excel_powerquery(action: "refresh", queryName: "Sales", loadDestination: "both")
```

For CLI users, use `pq-loadto` before `pq-refresh` to configure where data should load.

**pq-loadto** - Load connection-only query to worksheet

```powershell
excelcli pq-loadto <file.xlsx> <query-name> <sheet-name>
```

**pq-delete** - Delete Power Query

```powershell
excelcli pq-delete <file.xlsx> <query-name>
```

### Phase 1 Commands - Atomic Operations ✨ **NEW**

These commands provide atomic operations for cleaner Power Query workflows:

**pq-create** - Create new Power Query with atomic import + load

```powershell
excelcli pq-create <file.xlsx> <query-name> <mcode-file> [--destination worksheet|data-model|both|connection-only] [--target-sheet SheetName]

# Examples
excelcli pq-create "Sales.xlsx" "ImportData" "query.pq"                              # Default: load to worksheet
excelcli pq-create "Sales.xlsx" "DataModel" "query.pq" --destination data-model    # Load to Power Pivot
excelcli pq-create "Sales.xlsx" "Both" "query.pq" --destination both                # Load to both
excelcli pq-create "Sales.xlsx" "Connection" "query.pq" --destination connection-only  # Connection only
excelcli pq-create "Sales.xlsx" "Custom" "query.pq" --target-sheet "MySheet"        # Custom worksheet
```

Creates a new Power Query with M code and loads data in a single atomic operation. Replaces the two-step workflow of `pq-import` + `pq-loadto`.

**pq-update-mcode** - Update M code only (no refresh)

```powershell
excelcli pq-update-mcode <file.xlsx> <query-name> <mcode-file>

# Example
excelcli pq-update-mcode "Sales.xlsx" "ImportData" "updated-query.pq"
```

Updates only the M code without refreshing data. Use when you want to stage code changes before refreshing.

**pq-unload** - Convert query to connection-only

```powershell
excelcli pq-unload <file.xlsx> <query-name>

# Example
excelcli pq-unload "Sales.xlsx" "TempQuery"
```

Removes data from worksheet/data model while keeping the query definition. Inverse of `pq-loadto`.

**pq-update-and-refresh** - Update M code and refresh in one operation

```powershell
excelcli pq-update-and-refresh <file.xlsx> <query-name> <mcode-file>

# Example
excelcli pq-update-and-refresh "Sales.xlsx" "ImportData" "new-query.pq"
```

Updates M code and immediately refreshes data. Replaces the two-step workflow of `pq-update` + `pq-refresh`.

**pq-refresh-all** - Refresh all Power Queries

```powershell
excelcli pq-refresh-all <file.xlsx>

# Example
excelcli pq-refresh-all "Sales.xlsx"
```

Refreshes all Power Queries in the workbook in a single operation. Useful for batch data refreshes.

### Real-World Workflow Examples

**Scenario 1: Initial Project Setup**

Setting up a new data pipeline with Power Query:

```powershell
# Create workbook
excelcli create-empty "DataPipeline.xlsx"

# Import and load multiple queries atomically (vs old import + loadto pattern)
excelcli pq-create "DataPipeline.xlsx" "Sales" "queries/sales.pq" --destination both
excelcli pq-create "DataPipeline.xlsx" "Customers" "queries/customers.pq" --destination data-model
excelcli pq-create "DataPipeline.xlsx" "Products" "queries/products.pq" --destination worksheet

# Result: All queries created and data loaded in 3 commands (vs 6 with old workflow)
```

**Scenario 2: Iterative Development**

Developing and testing Power Query transformations:

```powershell
# Stage M code changes without waiting for refresh (fast iteration)
excelcli pq-update-mcode "DataPipeline.xlsx" "Sales" "queries/sales-v2.pq"
excelcli pq-update-mcode "DataPipeline.xlsx" "Customers" "queries/customers-v2.pq"

# Test changes manually in Excel UI, then refresh when ready
excelcli pq-refresh "DataPipeline.xlsx" "Sales"

# Or update and refresh together when code is finalized
excelcli pq-update-and-refresh "DataPipeline.xlsx" "Products" "queries/products-final.pq"
```

**Scenario 3: Production Data Refresh**

Deploying updates to production workbooks:

```powershell
# Atomic update + refresh for production (ensures code and data are in sync)
excelcli pq-update-and-refresh "Production.xlsx" "Sales" "prod/sales.pq"
excelcli pq-update-and-refresh "Production.xlsx" "Inventory" "prod/inventory.pq"

# Or batch refresh all queries without code changes
excelcli pq-refresh-all "Production.xlsx"

# Result: Production data current in single atomic operation per query
```

**Scenario 4: Cleanup and Optimization**

Managing queries for performance:

```powershell
# Remove data from unused queries to reduce file size
excelcli pq-unload "Analytics.xlsx" "OldQuery"
excelcli pq-unload "Analytics.xlsx" "TempTransform"

# Keep queries as connection-only for reference without loading data
excelcli pq-list "Analytics.xlsx"  # Verify queries still exist
```

**Comparison: Old vs New Workflows**

```powershell
# OLD WORKFLOW: Phase 3 commands (REMOVED)
# pq-import was removed - use pq-create instead
# pq-update was removed - use pq-update-mcode or pq-update-and-refresh instead

# NEW WORKFLOW (Phase 1 - atomic operations):
excelcli pq-create "file.xlsx" "Sales" "sales.pq"  # Create + load in one step
excelcli pq-update-and-refresh "file.xlsx" "Sales" "sales.pq"  # Update + refresh in one step

# Time savings: Atomic operations prevent intermediate states and reduce command count
```

## Sheet Commands (`sheet-*`)

Manage worksheet lifecycle (create, rename, copy, delete), tab colors, and visibility. For data operations, use `range-*` commands.

**sheet-list** - List all worksheets

```powershell
excelcli sheet-list <file.xlsx>
```

**sheet-create** - Create new worksheet

```powershell
excelcli sheet-create <file.xlsx> <sheet-name>
```

**sheet-rename** - Rename worksheet

```powershell
excelcli sheet-rename <file.xlsx> <old-name> <new-name>
```

**sheet-copy** - Copy worksheet

```powershell
excelcli sheet-copy <file.xlsx> <source-sheet> <new-sheet>
```

**sheet-delete** - Delete worksheet

```powershell
excelcli sheet-delete <file.xlsx> <sheet-name>
```

**sheet-set-tab-color** - Set worksheet tab color (RGB 0-255)

```powershell
excelcli sheet-set-tab-color <file.xlsx> <sheet-name> <red> <green> <blue>

# Examples
excelcli sheet-set-tab-color "Report.xlsx" "Sales" 255 0 0        # Red
excelcli sheet-set-tab-color "Report.xlsx" "Expenses" 0 255 0    # Green
excelcli sheet-set-tab-color "Report.xlsx" "Summary" 0 0 255     # Blue
```

**sheet-get-tab-color** - Get worksheet tab color

```powershell
excelcli sheet-get-tab-color <file.xlsx> <sheet-name>

# Example output
# Sheet: Sales
# Color: #FF0000 (Red: 255, Green: 0, Blue: 0)
```

**sheet-clear-tab-color** - Remove worksheet tab color

```powershell
excelcli sheet-clear-tab-color <file.xlsx> <sheet-name>
```

**sheet-set-visibility** - Set worksheet visibility level

```powershell
excelcli sheet-set-visibility <file.xlsx> <sheet-name> <visible|hidden|veryhidden>

# Examples
excelcli sheet-set-visibility "Report.xlsx" "Data" hidden          # User can unhide via UI
excelcli sheet-set-visibility "Report.xlsx" "Calculations" veryhidden  # Requires code to unhide
excelcli sheet-set-visibility "Report.xlsx" "Summary" visible      # Make visible
```

**sheet-get-visibility** - Get worksheet visibility level

```powershell
excelcli sheet-get-visibility <file.xlsx> <sheet-name>

# Example output
# Sheet: Data
# Visibility: Hidden
```

**sheet-show** - Show a hidden worksheet

```powershell
excelcli sheet-show <file.xlsx> <sheet-name>
```

**sheet-hide** - Hide a worksheet (user can unhide via UI)

```powershell
excelcli sheet-hide <file.xlsx> <sheet-name>
```

**sheet-very-hide** - Very hide a worksheet (requires code to unhide)

```powershell
excelcli sheet-very-hide <file.xlsx> <sheet-name>

# Example - protect calculations from users
excelcli sheet-very-hide "Model.xlsx" "Formulas"
```

## Range Commands (`range-*`)

Manage range data operations. Single cell = 1x1 range (e.g., "A1"). Named ranges: use empty sheet name "".

**range-get-values** - Read range values as CSV

```powershell
excelcli range-get-values <file.xlsx> <sheet-name> <range>

# Examples
excelcli range-get-values "Plan.xlsx" "Data" "A1:D10"  # Read range
excelcli range-get-values "Plan.xlsx" "Data" "A1"      # Read single cell (1x1 range)
excelcli range-get-values "Plan.xlsx" "" "SalesData"   # Read named range
```

**range-set-values** - Write CSV data to range

```powershell
excelcli range-set-values <file.xlsx> <sheet-name> <range> <data.csv>

# Examples
excelcli range-set-values "Plan.xlsx" "Data" "A1:C10" "values.csv"
excelcli range-set-values "Plan.xlsx" "" "SalesData" "sales.csv"  # Named range
```

**range-get-formulas** - Read range formulas as CSV

```powershell
excelcli range-get-formulas <file.xlsx> <sheet-name> <range>

# Example
excelcli range-get-formulas "Plan.xlsx" "Calc" "D1:D100"
```

**range-set-formulas** - Write CSV formulas to range

```powershell
excelcli range-set-formulas <file.xlsx> <sheet-name> <range> <formulas.csv>

# Example (formulas.csv contains: =SUM(A1:A10))
excelcli range-set-formulas "Plan.xlsx" "Calc" "D1" "total-formula.csv"
```

**range-clear-all** - Clear all content and formatting

```powershell
excelcli range-clear-all <file.xlsx> <sheet-name> <range>
```

**range-clear-contents** - Clear values/formulas, preserve formatting

```powershell
excelcli range-clear-contents <file.xlsx> <sheet-name> <range>
```

**range-clear-formats** - Clear formatting, preserve values/formulas

```powershell
excelcli range-clear-formats <file.xlsx> <sheet-name> <range>
```

### Number Formatting

**range-get-number-formats** - Get number format codes from range as CSV

```powershell
excelcli range-get-number-formats <file.xlsx> <sheet-name> <range>

# Example
excelcli range-get-number-formats "Sales.xlsx" "Sheet1" "A1:D10"
# Output (CSV): "$#,##0.00","0.00%","m/d/yyyy","General"
```

**range-set-number-format** - Apply uniform number format to range

```powershell
excelcli range-set-number-format <file.xlsx> <sheet-name> <range> <format-code>

# Examples
excelcli range-set-number-format "Sales.xlsx" "Sheet1" "D2:D100" "$#,##0.00"  # Currency
excelcli range-set-number-format "Sales.xlsx" "Sheet1" "E2:E100" "0.00%"      # Percentage
excelcli range-set-number-format "Sales.xlsx" "Sheet1" "A2:A100" "m/d/yyyy"   # Date
excelcli range-set-number-format "Sales.xlsx" "Sheet1" "B2:B100" "@"          # Text
```

### Visual Formatting

**range-format** - Apply visual formatting (font, fill, border, alignment)

```powershell
excelcli range-format <file.xlsx> <sheet-name> <range> [options]

# Font options
  --font-name NAME           Font family (e.g., Arial, Calibri)
  --font-size SIZE           Font size in points
  --bold                     Make text bold
  --italic                   Make text italic
  --underline                Underline text
  --font-color #RRGGBB       Font color in hex (e.g., #FF0000 for red)

# Fill options
  --fill-color #RRGGBB       Background color in hex

# Border options
  --border-style STYLE       Border style: Continuous, Dashed, Dotted, Double
  --border-weight WEIGHT     Border weight: Thin, Medium, Thick, Hairline
  --border-color #RRGGBB     Border color in hex

# Alignment options
  --h-align ALIGN            Horizontal: Left, Center, Right, Justify
  --v-align ALIGN            Vertical: Top, Center, Bottom
  --wrap-text                Enable text wrapping
  --orientation DEGREES      Text rotation (-90 to 90)

# Examples
excelcli range-format "Report.xlsx" "Sheet1" "A1:E1" --bold --font-size 12 --h-align Center  # Headers
excelcli range-format "Report.xlsx" "Sheet1" "D2:D100" --fill-color "#FFFF00"  # Yellow highlight
excelcli range-format "Report.xlsx" "Sheet1" "A1:E100" --border-style Continuous --border-weight Thin
```

### Data Validation

**range-validate** - Add data validation rules to range

```powershell
excelcli range-validate <file.xlsx> <sheet-name> <range> <type> <formula1> [formula2] [options]

# Validation types
  List            Dropdown list
  WholeNumber     Integer validation
  Decimal         Decimal number validation
  Date            Date validation
  Time            Time validation
  TextLength      Character length validation
  Custom          Custom formula validation

# Operators (for numeric/date validations)
  --operator OPERATOR        Between, NotBetween, Equal, NotEqual, Greater, Less, GreaterOrEqual, LessOrEqual

# Optional parameters
  --show-input               Show input message when cell selected
  --input-title TITLE        Input message title
  --input-message MSG        Input message text
  --error-title TITLE        Error alert title
  --error-message MSG        Error alert text
  --error-style STYLE        Stop, Warning, Information
  --ignore-blank             Ignore blank cells
  --show-dropdown            Show dropdown arrow

# Examples
excelcli range-validate "Data.xlsx" "Sheet1" "F2:F100" List "Active,Inactive,Pending"  # Dropdown
excelcli range-validate "Data.xlsx" "Sheet1" "E2:E100" WholeNumber "1" "999" --operator Between  # Number range
excelcli range-validate "Data.xlsx" "Sheet1" "C2:C100" TextLength "100" --operator LessOrEqual  # Max length
excelcli range-validate "Data.xlsx" "Sheet1" "A2:A100" Date "1/1/2025" --operator GreaterOrEqual  # Min date
```

### Migration from Sheet Commands

| Old Command | New Command | Notes |
|-------------|-------------|-------|
| `sheet-read` | `range-get-values` | Same functionality, works with any range |
| `sheet-write` | `range-set-values` | Specify range explicitly (e.g., "A1") |
| `sheet-clear` | `range-clear-all` or `range-clear-contents` | Choose based on whether to preserve formatting |
| `sheet-append` | *(not yet implemented)* | Use `range-set-values` with calculated range for now |

### CSV Conversion Behavior

- **Type Inference**: Numbers and booleans auto-detected, empty cells become null
- **Quote Escaping**: Values with commas, quotes, or newlines are automatically quoted
- **2D Arrays**: Core uses `List<List<object?>>`, CLI converts CSV ↔ 2D arrays for convenience

## Named Range Commands (`namedrange-*`)

Manage named ranges and parameters.

**namedrange-list** - List all named ranges

```powershell
excelcli namedrange-list <file.xlsx>
```

**namedrange-get** - Get named range value

```powershell
excelcli namedrange-get <file.xlsx> <namedrange-name>
```

**namedrange-set** - Set named range value

```powershell
excelcli namedrange-set <file.xlsx> <namedrange-name> <value>
```

**namedrange-update** - Update named range cell reference ✨ **NEW**

```powershell
excelcli namedrange-update <file.xlsx> <namedrange-name> <new-reference>
```

Updates the cell reference of a named range. Use `namedrange-set` to change the value, or `namedrange-update` to change which cell the parameter points to.

Example:
```powershell
# Change StartDate parameter from Sheet1!A1 to Config!B5
excelcli namedrange-update Sales.xlsx StartDate Config!B5
```

**namedrange-create** - Create named range

```powershell
excelcli namedrange-create <file.xlsx> <namedrange-name> <reference>
```

**namedrange-delete** - Delete named range

```powershell
excelcli namedrange-delete <file.xlsx> <namedrange-name>
```

## Connection Commands (`conn-*`)

Manage Excel connections (OLEDB, ODBC, Text, Web, etc.).

**conn-list** - List all connections in workbook

```powershell
excelcli conn-list <file.xlsx>
```

**conn-view** - Display connection details and connection string

```powershell
excelcli conn-view <file.xlsx> <connection-name>
```

**conn-import** - Import connection from ODC file

```powershell
excelcli conn-import <file.xlsx> <connection-name> <source.odc>
```

**conn-export** - Export connection to ODC file

```powershell
excelcli conn-export <file.xlsx> <connection-name> <output.odc>
```

**conn-update** - Update existing connection from ODC file

```powershell
excelcli conn-update <file.xlsx> <connection-name> <source.odc>
```

**conn-refresh** - Refresh connection data

```powershell
excelcli conn-refresh <file.xlsx> <connection-name>
```

**conn-delete** - Delete connection from workbook

```powershell
excelcli conn-delete <file.xlsx> <connection-name>
```

**conn-loadto** - Load connection-only connection to worksheet table

```powershell
excelcli conn-loadto <file.xlsx> <connection-name> <sheet-name>
```

**conn-properties** - Get connection properties (refresh settings, background query, etc.)

```powershell
excelcli conn-properties <file.xlsx> <connection-name>
```

**conn-set-properties** - Set connection properties

```powershell
excelcli conn-set-properties <file.xlsx> <connection-name> <property-json>
```

**conn-test** - Test connection connectivity

```powershell
excelcli conn-test <file.xlsx> <connection-name>
```

**Supported Connection Types:**
- **OLEDB** - OLE DB data sources (SQL Server, Access, etc.)
- **ODBC** - ODBC data sources
- **Text** - Text/CSV file imports
- **Web** - Web queries
- **DataFeed** - Data feed connections
- **Model** - Data model connections
- **Worksheet** - Worksheet connections

**⚠️ Important Notes:**
- Connection strings may contain sensitive data (passwords, credentials)
- Use `conn-export` carefully - ODC files may contain credentials
- Power Query connections use `pq-*` commands instead of `conn-*`

## Cell Commands (`cell-*`)

Manage individual cells.

**cell-get-value** - Get cell value

```powershell
excelcli cell-get-value <file.xlsx> <sheet-name> <cell>
```

**cell-set-value** - Set cell value

```powershell
excelcli cell-set-value <file.xlsx> <sheet-name> <cell> <value>
```

**cell-get-formula** - Get cell formula

```powershell
excelcli cell-get-formula <file.xlsx> <sheet-name> <cell>
```

**cell-set-formula** - Set cell formula

```powershell
excelcli cell-set-formula <file.xlsx> <sheet-name> <cell> <formula>
```

## Table Commands (`table-*`)

Manage Excel Tables (ListObjects) - structured data with auto-filtering, formatting, and Power Query integration.

### Overview

Excel Tables are structured ranges that provide:
- Automatic filtering and sorting
- Structured references in formulas
- Dynamic expansion when adding data
- Visual formatting with table styles
- Power Query integration via `Excel.CurrentWorkbook()`

### Lifecycle Commands

**table-list** - List all tables in workbook

```powershell
excelcli table-list <file.xlsx>
```

**table-create** - Create new table from range

```powershell
excelcli table-create <file.xlsx> <sheet-name> <table-name> <range> [hasHeaders] [tableStyle]

# Examples
excelcli table-create sales.xlsx Data SalesTable A1:E100
excelcli table-create sales.xlsx Data SalesTable A1:E100 true TableStyleMedium2
```

**table-info** - Get detailed table information

```powershell
excelcli table-info <file.xlsx> <table-name>
```

**table-rename** - Rename table

```powershell
excelcli table-rename <file.xlsx> <old-table-name> <new-table-name>
```

**table-delete** - Delete table (converts to range, preserves data)

```powershell
excelcli table-delete <file.xlsx> <table-name>
```

### Structure Commands

**table-resize** - Resize table to new range

```powershell
excelcli table-resize <file.xlsx> <table-name> <new-range>

# Example
excelcli table-resize sales.xlsx SalesTable A1:E150
```

**table-set-style** - Change table visual style

```powershell
excelcli table-set-style <file.xlsx> <table-name> <style-name>

# Example
excelcli table-set-style sales.xlsx SalesTable TableStyleDark1
```

**table-toggle-totals** - Show/hide totals row

```powershell
excelcli table-toggle-totals <file.xlsx> <table-name> <true|false>

# Example
excelcli table-toggle-totals sales.xlsx SalesTable true
```

**table-set-column-total** - Set total function for column

```powershell
excelcli table-set-column-total <file.xlsx> <table-name> <column-name> <function>

# Functions: sum, avg, count, max, min, stdev, var
# Example
excelcli table-set-column-total sales.xlsx SalesTable Amount sum
```

### Data Commands

**table-append** - Append rows to table

```powershell
excelcli table-append <file.xlsx> <table-name> <data.csv>

# Example
excelcli table-append sales.xlsx SalesTable new-rows.csv
```

### Filter Commands ✨ **NEW**

**table-apply-filter** - Filter table column by criteria

```powershell
excelcli table-apply-filter <file.xlsx> <table-name> <column-name> <criteria>

# Criteria operators: >value, <value, =value, >=value, <=value, <>value
# Examples
excelcli table-apply-filter sales.xlsx SalesTable Amount ">100"
excelcli table-apply-filter sales.xlsx SalesTable Status "=Active"
excelcli table-apply-filter sales.xlsx SalesTable Region "<>North"
```

**table-apply-filter-values** - Filter table column by specific values

```powershell
excelcli table-apply-filter-values <file.xlsx> <table-name> <column-name> <value1,value2,...>

# Example - show only North, South, East regions
excelcli table-apply-filter-values sales.xlsx SalesTable Region "North,South,East"
```

**table-clear-filters** - Remove all filters from table

```powershell
excelcli table-clear-filters <file.xlsx> <table-name>

# Example
excelcli table-clear-filters sales.xlsx SalesTable
```

**table-get-filters** - Get current filter state

```powershell
excelcli table-get-filters <file.xlsx> <table-name>

# Displays table of filtered columns with criteria and values
```

### Column Commands ✨ **NEW**

**table-add-column** - Add new column to table

```powershell
excelcli table-add-column <file.xlsx> <table-name> <column-name> [position]

# Examples
excelcli table-add-column sales.xlsx SalesTable NewColumn
excelcli table-add-column sales.xlsx SalesTable NewColumn 2
```

**table-remove-column** - Remove column from table

```powershell
excelcli table-remove-column <file.xlsx> <table-name> <column-name>

# Example
excelcli table-remove-column sales.xlsx SalesTable OldColumn
```

**table-rename-column** - Rename table column

```powershell
excelcli table-rename-column <file.xlsx> <table-name> <old-column-name> <new-column-name>

# Example
excelcli table-rename-column sales.xlsx SalesTable OldName NewName
```

### Data Model Integration

**table-add-to-datamodel** - Add table to Power Pivot Data Model

```powershell
excelcli table-add-to-datamodel <file.xlsx> <table-name>

# Example
excelcli table-add-to-datamodel sales.xlsx SalesTable
```

### Structured Reference Operations ✨ **NEW**

**table-get-structured-reference** - Get structured reference formula for table region

```powershell
excelcli table-get-structured-reference <file.xlsx> <table-name> <region> [column-name]

# Regions: All, Data, Headers, Totals, ThisRow
# Examples
excelcli table-get-structured-reference sales.xlsx SalesTable Data
# Returns: SalesTable[#Data] and range address $A$2:$D$100

excelcli table-get-structured-reference sales.xlsx SalesTable Data Amount
# Returns: SalesTable[[Amount]] and range address $D$2:$D$100

excelcli table-get-structured-reference sales.xlsx SalesTable Headers
# Returns: SalesTable[#Headers] and range address $A$1:$D$1
```

**Workflow Hints:**
- Use with RangeCommands: Get the range address, then use `range-get-values` to read data
- Use in formulas: Copy the structured reference for use in Excel formulas
- Table regions: All (entire table), Data (rows only), Headers (header row), Totals (totals row)

### Sort Operations ✨ **NEW**

**table-sort** - Sort table by single column

```powershell
excelcli table-sort <file.xlsx> <table-name> <column-name> [asc|desc]

# Examples
excelcli table-sort sales.xlsx SalesTable Amount desc
excelcli table-sort sales.xlsx SalesTable Date asc
```

**table-sort-multi** - Sort table by multiple columns (max 3 levels)

```powershell
excelcli table-sort-multi <file.xlsx> <table-name> <column1:asc> <column2:desc> [column3:asc]

# Examples
excelcli table-sort-multi sales.xlsx SalesTable Region:asc Amount:desc
excelcli table-sort-multi sales.xlsx SalesTable Year:desc Quarter:desc Amount:desc
```

**Workflow Hints:**
- Single column sort: Simple ascending/descending sort
- Multi-column sort: Excel supports max 3 sort levels
- Table structure preserved: Headers and totals row maintained

## PivotTable Commands (`pivot-*`)

Manage Excel PivotTables for interactive data analysis and summarization. Create PivotTables from ranges, Excel Tables, or Power Pivot Data Model tables.

### Overview

PivotTables provide powerful data analysis capabilities:
- Dynamic data summarization with drag-and-drop field configuration
- Multiple aggregation functions (Sum, Count, Average, Max, Min, etc.)
- Row, Column, Value, and Filter field areas
- Automatic refresh from source data
- Integration with Power Pivot Data Model for large datasets

### Creation Commands

**pivot-create-from-range** - Create PivotTable from range

```powershell
excelcli pivot-create-from-range <file.xlsx> <source-sheet> <source-range> <dest-sheet> <dest-cell> <pivot-name>
```

Example:
```powershell
excelcli pivot-create-from-range sales.xlsx Data A1:D100 Analysis A1 SalesPivot
```

Creates a PivotTable from a data range with headers. The range must include at least 2 rows (headers + data).

**pivot-create-from-datamodel** - Create PivotTable from Power Pivot Data Model table

```powershell
excelcli pivot-create-from-datamodel <file.xlsx> <datamodel-table-name> <dest-sheet> <dest-cell> <pivot-name>
```

Example:
```powershell
excelcli pivot-create-from-datamodel sales.xlsx ConsumptionMilestones Analysis A1 MilestonesPivot
```

Creates a PivotTable from a table in the Power Pivot Data Model. This enables:
- Analysis of large datasets (millions of rows)
- Use of DAX measures in PivotTables
- Relationships between multiple tables
- Professional BI solutions integrated with Power BI

**Use Cases:**
- Automating analytical dashboards with Data Model tables
- Creating PivotTables from imported data models
- Integrating with Azure consumption planning tools (e.g., CP Toolkit)

**Requirements:**
- Workbook must contain a Power Pivot Data Model
- Table must exist in the Data Model (use `dm-list-tables` to verify)

### Management Commands

**pivot-list** - List all PivotTables

```powershell
excelcli pivot-list <file.xlsx>
```

Displays all PivotTables in the workbook with details:
- PivotTable name and sheet location
- Source data reference
- Field counts (Row, Column, Value fields)

**pivot-add-row-field** - Add field to Row area

```powershell
excelcli pivot-add-row-field <file.xlsx> <pivot-name> <field-name> [position]
```

Example:
```powershell
excelcli pivot-add-row-field sales.xlsx SalesPivot Region
excelcli pivot-add-row-field sales.xlsx SalesPivot Product 2
```

**pivot-add-value-field** - Add field to Values area

```powershell
excelcli pivot-add-value-field <file.xlsx> <pivot-name> <field-name> [function] [custom-name]
```

Example:
```powershell
excelcli pivot-add-value-field sales.xlsx SalesPivot Amount Sum "Total Sales"
excelcli pivot-add-value-field sales.xlsx SalesPivot Quantity Count
```

Supported aggregation functions: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP

**pivot-refresh** - Refresh PivotTable data

```powershell
excelcli pivot-refresh <file.xlsx> <pivot-name>
```

Refreshes the PivotTable to reflect changes in source data.

**Workflow Hints:**
- Create PivotTable from Data Model for large datasets and DAX measures
- Use `dm-list-tables` to find available Data Model tables
- Add Row fields for grouping, Value fields for calculations
- Refresh PivotTables after source data changes

## VBA VBA Commands (`vba-*`)

**⚠️ VBA commands require macro-enabled (.xlsm) files!**

Manage VBA scripts and macros in macro-enabled Excel workbooks.

**vba-list** - List all VBA modules and procedures

```powershell
excelcli vba-list <file.xlsm>
```

**vba-view** - View VBA module code ✨ **NEW**

```powershell
excelcli vba-view <file.xlsm> <module-name>
```

Displays the complete VBA code for a module without exporting to a file. Shows module type, line count, procedures, and full source code.

Example:
```powershell
# View the DataProcessor module code
excelcli vba-view Report.xlsm DataProcessor
```

**vba-export** - Export VBA module to file

```powershell
excelcli vba-export <file.xlsm> <module-name> <output.vba>
```

**vba-import** - Import VBA module from file

```powershell
excelcli vba-import <file.xlsm> <module-name> <source.vba>
```

**vba-update** - Update existing VBA module

```powershell
excelcli vba-update <file.xlsm> <module-name> <source.vba>
```

**vba-run** - Execute VBA macro with parameters

```powershell
excelcli vba-run <file.xlsm> <macro-name> [param1] [param2] ...

# Examples
excelcli vba-run "Report.xlsm" "ProcessData"
excelcli vba-run "Analysis.xlsm" "CalculateTotal" "Sheet1" "A1:C10"
```

**vba-delete** - Remove VBA module

```powershell
excelcli vba-delete <file.xlsm> <module-name>
```

## Data Model Commands (`dm-*`)

Manage Excel Data Model (Power Pivot) - tables, DAX measures, and relationships. The Data Model provides enterprise-grade data analysis capabilities built into Excel.

### Overview

The Data Model is Excel's in-memory analytical database (formerly known as Power Pivot). It supports:
- Large datasets (millions of rows)
- Relationships between tables
- DAX (Data Analysis Expressions) calculated measures
- DAX calculated columns
- Advanced analytics and aggregations

**Note:** CREATE and UPDATE operations for measures, relationships, and calculated columns use the Analysis Services Tabular Object Model (TOM) API. READ operations use Excel COM API where available.

### Commands

**dm-list-tables** - List all Data Model tables

```powershell
excelcli dm-list-tables <file.xlsx>
```

Displays all tables loaded into the Data Model with record counts and source information.

**dm-list-measures** - List all DAX measures

```powershell
excelcli dm-list-measures <file.xlsx>
```

Shows all calculated measures with their DAX formulas and associated tables.

**dm-view-measure** - View specific measure formula

```powershell
excelcli dm-view-measure <file.xlsx> <measure-name>
```

Displays the complete DAX formula for a specific measure.

**dm-export-measure** - Export measure to DAX file

```powershell
excelcli dm-export-measure <file.xlsx> <measure-name> <output.dax>
```

Exports a measure's DAX formula to a file for version control or documentation.

**dm-list-relationships** - List all table relationships

```powershell
excelcli dm-list-relationships <file.xlsx>
```

Shows all relationships between Data Model tables, including direction and active status.

**dm-refresh** - Refresh all Data Model tables

```powershell
excelcli dm-refresh <file.xlsx>
```

Refreshes all tables in the Data Model, loading latest data from source connections.

**dm-delete-measure** - Delete a DAX measure

```powershell
excelcli dm-delete-measure <file.xlsx> <measure-name>
```

Permanently removes a measure from the Data Model. Changes are saved to the workbook.

**dm-delete-relationship** - Delete a table relationship

```powershell
excelcli dm-delete-relationship <file.xlsx> <from-table> <from-column> <to-table> <to-column>
```

Removes a relationship between two tables in the Data Model. Changes are saved to the workbook.

### Phase 2: Discovery & CREATE/UPDATE Operations ✨ **NEW**

**dm-list-columns** - List columns in a Data Model table

```powershell
excelcli dm-list-columns <file.xlsx> <table-name>
```

Lists all columns in a table with their data types and calculated status.

Example:
```powershell
excelcli dm-list-columns "sales-model.xlsx" "Sales"
# Output: Column Name, Data Type, Calculated
#         SalesID, Integer, No
#         Amount, Currency, No
#         TotalWithTax, Currency, Yes
```

**dm-view-table** - View table details with columns and measures

```powershell
excelcli dm-view-table <file.xlsx> <table-name>
```

Shows complete table information including source, record count, refresh date, columns, and measure count.

Example:
```powershell
excelcli dm-view-table "sales-model.xlsx" "Sales"
# Output: Table Name, Source, Record Count, Columns (detailed list), Measure Count
```

**dm-get-model-info** - Get Data Model overview

```powershell
excelcli dm-get-model-info <file.xlsx>
```

Shows Data Model statistics: table count, measure count, relationship count, total rows, and table names.

Example:
```powershell
excelcli dm-get-model-info "sales-model.xlsx"
# Output: Tables: 3, Measures: 5, Relationships: 2, Total Rows: 15,234
```

**dm-create-measure** - Create DAX measure

```powershell
excelcli dm-create-measure <file.xlsx> <table-name> <measure-name> <dax-formula> [format-type] [description]
```

Creates a new DAX measure in the specified table. Format types: `Currency`, `Decimal`, `Percentage`, `General`.

Examples:
```powershell
# Create measure with currency format
excelcli dm-create-measure "sales.xlsx" "Sales" "TotalRevenue" "SUM(Sales[Amount])" "Currency" "Total sales revenue"

# Create percentage measure
excelcli dm-create-measure "sales.xlsx" "Sales" "GrowthRate" "DIVIDE(SUM(Sales[CurrentYear]), SUM(Sales[PriorYear])) - 1" "Percentage"

# Simple measure (no format)
excelcli dm-create-measure "sales.xlsx" "Products" "ProductCount" "COUNTROWS(Products)"
```

**dm-update-measure** - Update existing measure

```powershell
excelcli dm-update-measure <file.xlsx> <measure-name> [dax-formula] [format-type] [description]
```

Updates an existing measure. At least one optional parameter must be provided.

Examples:
```powershell
# Update formula only
excelcli dm-update-measure "sales.xlsx" "TotalRevenue" "CALCULATE(SUM(Sales[Amount]))"

# Update format only
excelcli dm-update-measure "sales.xlsx" "TotalRevenue" "" "Decimal"

# Update description only
excelcli dm-update-measure "sales.xlsx" "TotalRevenue" "" "" "Updated revenue calculation"

# Update multiple properties
excelcli dm-update-measure "sales.xlsx" "GrowthRate" "DIVIDE([CurrentYearSales], [PriorYearSales]) - 1" "Percentage" "Year-over-year growth"
```

**dm-create-relationship** - Create table relationship

```powershell
excelcli dm-create-relationship <file.xlsx> <from-table> <from-column> <to-table> <to-column> [active:true|false]
```

Creates a relationship between two tables. Default: active=true.

Examples:
```powershell
# Create active relationship
excelcli dm-create-relationship "sales.xlsx" "Sales" "CustomerID" "Customers" "ID"

# Create inactive relationship
excelcli dm-create-relationship "sales.xlsx" "Sales" "AlternateCustomerID" "Customers" "ID" "false"
```

**dm-update-relationship** - Update relationship active status

```powershell
excelcli dm-update-relationship <file.xlsx> <from-table> <from-column> <to-table> <to-column> <active:true|false>
```

Toggles a relationship's active status. Only one relationship between two tables can be active at a time.

Examples:
```powershell
# Activate relationship
excelcli dm-update-relationship "sales.xlsx" "Sales" "CustomerID" "Customers" "ID" "true"

# Deactivate relationship
excelcli dm-update-relationship "sales.xlsx" "Sales" "CustomerID" "Customers" "ID" "false"
```

### Usage Examples

```powershell
# Discovery - Explore Data Model structure
excelcli dm-get-model-info "sales-analysis.xlsx"
excelcli dm-list-tables "sales-analysis.xlsx"
excelcli dm-view-table "sales-analysis.xlsx" "Sales"
excelcli dm-list-columns "sales-analysis.xlsx" "Sales"

# Measures - Create and manage DAX calculations
excelcli dm-create-measure "sales.xlsx" "Sales" "TotalRevenue" "SUM(Sales[Amount])" "Currency"
excelcli dm-list-measures "sales-analysis.xlsx"
excelcli dm-view-measure "sales-analysis.xlsx" "TotalRevenue"
excelcli dm-update-measure "sales.xlsx" "TotalRevenue" "CALCULATE(SUM(Sales[Amount]))" "Currency" "Updated formula"

# Relationships - Connect tables
excelcli dm-create-relationship "sales.xlsx" "Sales" "CustomerID" "Customers" "ID"
excelcli dm-list-relationships "sales-analysis.xlsx"
excelcli dm-update-relationship "sales.xlsx" "Sales" "CustomerID" "Customers" "ID" "false"

# Export measure for version control
excelcli dm-export-measure "sales-analysis.xlsx" "Total Sales" "measures/total-sales.dax"

# Refresh data
excelcli dm-refresh "sales-analysis.xlsx"

# Delete operations
excelcli dm-delete-measure "sales-analysis.xlsx" "Old Measure"
excelcli dm-delete-relationship "sales-analysis.xlsx" "Sales" "CustomerID" "Customers" "ID"
```

### CRUD Operations Status

| Operation | Status | Technology |
|-----------|--------|------------|
| **CREATE** | ✅ Available | Excel COM API |
| **READ** | ✅ Available | Excel COM API |
| **UPDATE** | ✅ Available | Excel COM API |
| **DELETE** | ✅ Available | Excel COM API |

**Available Capabilities:**
- ✅ Create DAX measures with format types and descriptions (Excel COM API)
- ✅ Update existing measure properties (formula, format, description) (Excel COM API)
- ✅ Create table relationships with active/inactive flag (Excel COM API)
- ✅ Update relationship active status (toggle on/off) (Excel COM API)
- ✅ Discover model structure (tables, columns, measures, relationships)

**Note:** Calculated columns are not supported via automation. Use Excel UI to create calculated columns.

**Advanced Operations (Future):**
- Hierarchies
- Perspectives
- KPIs
- Advanced formatting options

## VBA Trust Configuration

VBA operations require **"Trust access to the VBA project object model"** to be enabled in Excel settings. This is a one-time manual setup for security reasons.

### How to Enable VBA Trust

1. Open Excel
2. Go to **File → Options → Trust Center**
3. Click **"Trust Center Settings"**
4. Select **"Macro Settings"**
5. Check **"✓ Trust access to the VBA project object model"**
6. Click **OK** twice to save settings

After enabling this setting, VBA operations will work automatically. If VBA trust is not enabled, commands will display detailed instructions.

**Security Note:** ExcelMcp never modifies security settings automatically. Users must explicitly enable VBA trust through Excel's settings to maintain security control.

For more information, see [Microsoft's documentation on macro security](https://support.microsoft.com/office/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6).

## Power Query Privacy Levels

When Power Query combines data from multiple sources, Excel requires a privacy level to be specified for security. ExcelMcp provides explicit control through the `--privacy-level` parameter.

### Privacy Level Options

- **None** - Ignores privacy levels, allows combining any data sources (least secure)
- **Private** - Prevents sharing data with other sources (most secure, recommended for sensitive data)
- **Organizational** - Data can be shared within organization (recommended for internal data)
- **Public** - Publicly available data sources (appropriate for public APIs)

### Using Privacy Levels

```powershell
# Specify privacy level explicitly
excelcli pq-import data.xlsx "WebData" query.pq --privacy-level Private

# Set default via environment variable (useful for automation)
$env:EXCEL_DEFAULT_PRIVACY_LEVEL = "Private"
excelcli pq-import data.xlsx "WebData" query.pq
```

If a privacy level is needed but not specified, the command will display:
- Existing privacy levels in the workbook
- Recommended privacy level based on your queries
- Clear instructions on how to proceed

**Security Note:** ExcelMcp never applies privacy levels automatically. Users must explicitly choose the appropriate level for their data security requirements.

## File Format Support

- **`.xlsx`** - Standard Excel workbooks (Power Query, worksheets, parameters)
- **`.xlsm`** - Macro-enabled workbooks (includes VBA script support)

Use `create-empty` with `.xlsm` extension to create macro-enabled workbooks:

```powershell
excelcli create-empty "macros.xlsm"  # Creates macro-enabled workbook
excelcli create-empty "data.xlsx"    # Creates standard workbook
```
