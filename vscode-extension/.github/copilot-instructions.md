# Excel MCP Server - Usage Instructions

> **How to use the Excel MCP Server tools to automate Microsoft Excel**

## Prerequisites

- **Windows OS** - Excel COM automation requires Windows
- **Microsoft Excel 2016 or later** - Must be installed
- **File must be CLOSED in Excel** - ExcelMcp requires exclusive file access

---

## Available Tools (12 tools, 172 operations)

| Tool | Purpose |
|------|---------|
| `excel_file` | Open/close sessions, create workbooks |
| `excel_powerquery` | Import external data, M code management |
| `excel_datamodel` | DAX measures, relationships, Power Pivot |
| `excel_table` | Excel Tables with filtering and sorting |
| `excel_pivottable` | PivotTables from ranges/tables/data model |
| `excel_chart` | Create and configure charts |
| `excel_range` | Cell values, formulas, formatting |
| `excel_worksheet` | Sheet lifecycle (create, delete, rename) |
| `excel_namedrange` | Named ranges for parameters |
| `excel_connection` | Manage OLEDB/ODBC connections |
| `excel_vba` | VBA modules (.xlsm files only) |
| `excel_conditionalformat` | Conditional formatting rules |

---

## Critical Rules

### Session Management
- **ALWAYS** start with `excel_file(action: 'open')` before any operations
- **KEEP session open** across multiple operations
- **ONLY close** with `excel_file(action: 'close')` when all operations complete
- Ask user "Should I close the session now?" if unclear

### File Access
- File MUST be closed in Excel UI before automation
- Tell user: "Please close the file in Excel before running automation"

---

## Tool Selection Guide

| Need | Use | NOT |
|------|-----|-----|
| External data (CSV, databases, APIs) | `excel_powerquery` | `excel_table` |
| DAX measures / Data Model | `excel_datamodel` | `excel_range` |
| Worksheet formulas | `excel_range` | `excel_datamodel` |
| Convert range to table | `excel_table` | - |
| Sheet lifecycle | `excel_worksheet` | - |
| Named ranges | `excel_namedrange` | - |
| VBA macros | `excel_vba` (.xlsm only) | - |

---

## Power Query (excel_powerquery)

**Actions**: list, view, create, update, refresh, load-to, unload, delete

**Load destinations** (critical for DAX):
- `worksheet` - Users see data, NO DAX capability (default)
- `data-model` - Ready for DAX, users DON'T see data
- `both` - Users see data AND DAX works

**Common patterns**:
```
# Import CSV to worksheet
excel_powerquery(action: 'create', mCode: 'let Source = Csv.Document(...) in Source', loadDestination: 'worksheet')

# Import for DAX analysis
excel_powerquery(action: 'create', mCode: '...', loadDestination: 'data-model')
```

**Mistakes to avoid**:
- Using `create` on existing query → use `update` instead
- Omitting `loadDestination` when DAX is needed

---

## Data Model & DAX (excel_datamodel)

**Actions**: list-tables, read-table, list-columns, create-measure, update-measure, delete-measure, list-measures, list-relationships, create-relationship, delete-relationship

**Prerequisites**: Data must be loaded with `loadDestination: 'data-model'` or `'both'`

**DAX syntax** (NOT worksheet formulas):
```
# Create measure
excel_datamodel(action: 'create-measure', tableName: 'Sales', measureName: 'Total Revenue', formula: 'SUM(Sales[Amount])')
```

---

## Range Operations (excel_range)

**Actions**: get-values, set-values, get-formulas, set-formulas, format-range, clear-all, clear-contents, copy, sort, find, replace, validate-range, add-hyperlink, merge-cells, autofit-columns, and more (42 total)

**Quirks**:
- Single cell returns `[[value]]` (2D array), NOT scalar
- Named ranges: use `sheetName: ''` (empty string)

**Format vs Style**:
- Use `set-style` for built-in Excel styles (faster, theme-aware)
- Use `format-range` only for custom brand colors

---

## Tables (excel_table)

**Actions**: create, list, get, resize, rename, delete, add-column, remove-column, append-rows, apply-filter, clear-filter, sort, get-data, add-to-datamodel, and more (24 total)

**When to use**: Structured data with headers, AutoFilter, structured references

---

## PivotTables (excel_pivottable)

**Actions**: create, list, get, add-field, remove-field, set-aggregation, apply-filter, sort-field, refresh, get-data, delete, and more (25 total)

**Create from different sources**:
```
# From range
excel_pivottable(action: 'create', sourceType: 'range', sourceRange: 'Sheet1!A1:D100')

# From table
excel_pivottable(action: 'create', sourceType: 'table', tableName: 'SalesData')

# From data model
excel_pivottable(action: 'create', sourceType: 'datamodel')
```

---

## Common Mistakes

| ❌ Wrong | ✅ Correct |
|----------|-----------|
| `excel_table(action: 'create')` for CSV import | `excel_powerquery(action: 'create', loadDestination: 'worksheet')` |
| `excel_powerquery(action: 'create')` for DAX | `excel_powerquery(action: 'create', loadDestination: 'data-model')` |
| `excel_namedrange(action: 'create')` called 5 times | `excel_namedrange(action: 'create-bulk')` with array |
| `excel_range` for DAX formulas | `excel_datamodel(action: 'create-measure')` |
| Working on file open in Excel | Ask user to close file first |
| Closing session mid-workflow | Keep session open until complete |

---

## Error Handling

**"File is already open"** → Tell user to close Excel file first

**"Value does not fall within expected range"** → Usually invalid range address or unsupported operation

**"Query 'X' already exists"** → Use `update` action instead of `create`

**Refresh timeout** → Ask user to run refresh manually in Excel

---

## Example Workflows

### Import CSV and Create PivotTable
1. `excel_file(action: 'open', filePath: 'workbook.xlsx')`
2. `excel_powerquery(action: 'create', mCode: '...', loadDestination: 'worksheet')`
3. `excel_table(action: 'list')` - verify table created
4. `excel_pivottable(action: 'create', sourceType: 'table', tableName: '...')`
5. `excel_pivottable(action: 'add-field', fieldName: 'Region', area: 'Row')`
6. `excel_file(action: 'close', save: true)`

### Build DAX Measure
1. `excel_file(action: 'open', filePath: 'workbook.xlsx')`
2. `excel_powerquery(action: 'create', loadDestination: 'data-model', mCode: '...')`
3. `excel_datamodel(action: 'create-measure', tableName: 'Sales', measureName: 'YoY Growth', formula: '...')`
4. `excel_file(action: 'close', save: true)`

---

## References

- [Complete Feature Reference](https://github.com/sbroenne/mcp-server-excel/blob/main/FEATURES.md)
- [Main Documentation](https://github.com/sbroenne/mcp-server-excel)
