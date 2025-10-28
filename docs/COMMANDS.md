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

### Usage Examples

```powershell
# View Data Model structure
excelcli dm-list-tables "sales-analysis.xlsx"
excelcli dm-list-measures "sales-analysis.xlsx"
excelcli dm-list-relationships "sales-analysis.xlsx"

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
| **CREATE** | Phase 4 (Future) | Requires TOM API |
| **READ** | ✅ Available | Excel COM API |
| **UPDATE** | Phase 4 (Future) | Requires TOM API |
| **DELETE** | ✅ Available | Excel COM API |

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
