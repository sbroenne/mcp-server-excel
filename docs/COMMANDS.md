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
  • Run the VBA procedure with script-run
  • Export the module for version control with script-export
  • Test the module functionality
```

### New Features
- `--connection-only` flag for `pq-import` to create queries without loading data to worksheets
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

**pq-import** - Create or import query from file

```powershell
excelcli pq-import <file.xlsx> <query-name> <source.pq> [--privacy-level <None|Private|Organizational|Public>]
```

Import a Power Query from an M code file. If the query combines data from multiple sources and privacy level is not specified, you'll receive guidance on which privacy level to choose.

**pq-export** - Export query to file

```powershell
excelcli pq-export <file.xlsx> <query-name> <output.pq>
```

**pq-update** - Update existing query from file

```powershell
excelcli pq-update <file.xlsx> <query-name> <code.pq> [--privacy-level <None|Private|Organizational|Public>]
```

Update an existing Power Query with new M code. If the query combines data from multiple sources and privacy level is not specified, you'll receive guidance on which privacy level to choose.

**pq-refresh** - Refresh query data

```powershell
excelcli pq-refresh <file.xlsx> <query-name>
```

**pq-loadto** - Load connection-only query to worksheet

```powershell
excelcli pq-loadto <file.xlsx> <query-name> <sheet-name>
```

**pq-delete** - Delete Power Query

```powershell
excelcli pq-delete <file.xlsx> <query-name>
```

## Sheet Commands (`sheet-*`)

Manage worksheet lifecycle (create, rename, copy, delete). For data operations, use `range-*` commands.

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

## Parameter Commands (`param-*`)

Manage named ranges and parameters.

**param-list** - List all named ranges

```powershell
excelcli param-list <file.xlsx>
```

**param-get** - Get named range value

```powershell
excelcli param-get <file.xlsx> <param-name>
```

**param-set** - Set named range value

```powershell
excelcli param-set <file.xlsx> <param-name> <value>
```

**param-update** - Update named range cell reference ✨ **NEW**

```powershell
excelcli param-update <file.xlsx> <param-name> <new-reference>
```

Updates the cell reference of a named range. Use `param-set` to change the value, or `param-update` to change which cell the parameter points to.

Example:
```powershell
# Change StartDate parameter from Sheet1!A1 to Config!B5
excelcli param-update Sales.xlsx StartDate Config!B5
```

**param-create** - Create named range

```powershell
excelcli param-create <file.xlsx> <param-name> <reference>
```

**param-delete** - Delete named range

```powershell
excelcli param-delete <file.xlsx> <param-name>
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

## VBA Script Commands (`script-*`)

**⚠️ VBA commands require macro-enabled (.xlsm) files!**

Manage VBA scripts and macros in macro-enabled Excel workbooks.

**script-list** - List all VBA modules and procedures

```powershell
excelcli script-list <file.xlsm>
```

**script-view** - View VBA module code ✨ **NEW**

```powershell
excelcli script-view <file.xlsm> <module-name>
```

Displays the complete VBA code for a module without exporting to a file. Shows module type, line count, procedures, and full source code.

Example:
```powershell
# View the DataProcessor module code
excelcli script-view Report.xlsm DataProcessor
```

**script-export** - Export VBA module to file

```powershell
excelcli script-export <file.xlsm> <module-name> <output.vba>
```

**script-import** - Import VBA module from file

```powershell
excelcli script-import <file.xlsm> <module-name> <source.vba>
```

**script-update** - Update existing VBA module

```powershell
excelcli script-update <file.xlsm> <module-name> <source.vba>
```

**script-run** - Execute VBA macro with parameters

```powershell
excelcli script-run <file.xlsm> <macro-name> [param1] [param2] ...

# Examples
excelcli script-run "Report.xlsm" "ProcessData"
excelcli script-run "Analysis.xlsm" "CalculateTotal" "Sheet1" "A1:C10"
```

**script-delete** - Remove VBA module

```powershell
excelcli script-delete <file.xlsm> <module-name>
```

## Data Model Commands (`dm-*`)

Manage Excel Data Model (Power Pivot) - tables, DAX measures, and relationships. The Data Model provides enterprise-grade data analysis capabilities built into Excel.

### Overview

The Data Model is Excel's in-memory analytical database (formerly known as Power Pivot). It supports:
- Large datasets (millions of rows)
- Relationships between tables
- DAX (Data Analysis Expressions) calculated measures
- Advanced analytics and aggregations

**Note:** CREATE and UPDATE operations for measures and relationships require the Analysis Services Tabular Object Model (TOM) API, which is planned for a future phase. Current operations support READ and DELETE via Excel COM API.

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
| **CREATE** | ✅ Available (Phase 2) | Excel COM API |
| **READ** | ✅ Available | Excel COM API |
| **UPDATE** | ✅ Available (Phase 2) | Excel COM API |
| **DELETE** | ✅ Available | Excel COM API |

**Phase 2 Capabilities:**
- ✅ Create DAX measures with format types and descriptions
- ✅ Update existing measure properties (formula, format, description)
- ✅ Create table relationships with active/inactive flag
- ✅ Update relationship active status (toggle on/off)
- ✅ Discover model structure (tables, columns, measures, relationships)

**Advanced Operations (Future Phase 4 - TOM API):**
- Calculated columns
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
