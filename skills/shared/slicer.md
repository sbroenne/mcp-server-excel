````markdown
# slicer - Server Quirks

**Slicer Types**:

Two distinct slicer types exist:
- **PivotTable Slicers**: Filter PivotTables (can control multiple PivotTables)
- **Table Slicers**: Filter Excel Tables (single table only)

**Actions**:

| Action | Description | Required Parameters |
|--------|-------------|---------------------|
| `create-slicer` | Create PivotTable slicer | pivotTableName, fieldName |
| `list-slicers` | List all PivotTable slicers | (none) |
| `set-slicer-selection` | Set PivotTable slicer filter | slicerName, selectedItems |
| `delete-slicer` | Delete PivotTable slicer | slicerName |
| `create-table-slicer` | Create Table slicer | tableName, columnName |
| `list-table-slicers` | List all Table slicers | (none) |
| `set-table-slicer-selection` | Set Table slicer filter | slicerName, selectedItems |
| `delete-table-slicer` | Delete Table slicer | slicerName |

**CRITICAL: Required Parameters** - The "Required Parameters" column above is strict. Missing any required parameter will cause an error. Pay special attention to `pivotTableName` for PivotTable slicers and `slicerName` for selection/deletion operations.

**Naming Convention**:

- If `slicerName` not provided, auto-generates `{FieldName}Slicer` or `{ColumnName}Slicer`
- Slicer names must be unique within workbook
- Use `list-slicers` or `list-table-slicers` to check existing names

**Selection Behavior**:

- `selectedItems` is a list of strings: `["Value1", "Value2"]`
- Empty list `[]` clears all filters (shows all items)
- Values must match exactly (case-sensitive)
- Invalid values are silently ignored

**CLI: JSON Array Quoting** (important for `--selected-items`):

The `--selected-items` parameter requires a JSON array. Use proper shell escaping:

```powershell
# PowerShell: use single quotes around the JSON, double quotes inside
--selected-items '["West","East"]'

# Or escape inner quotes with backtick
--selected-items "[`"West`",`"East`"]"

# Clear filter (show all items)
--selected-items '[]'
```

**Positioning**:

- `destinationSheet` specifies which worksheet hosts the slicer
- `position` is a cell address for top-left corner (e.g., `'E1'`, `'G5'`)
- The slicer's top-left corner aligns to the specified cell
- Default position if not specified: Excel chooses

**Common Mistakes**:

- Creating slicer for field not in PivotTable → Error
- Creating table slicer for column not in table → Error
- Setting selection with wrong case → Values ignored (filter shows nothing)
- Deleting slicer that doesn't exist → Error

**Best Practices**:

1. Call `list-slicers` before creating to avoid name conflicts
2. Use `list-slicers` to get exact slicer names for selection/deletion
3. Multi-PivotTable filtering: Create one slicer, connect to multiple PivotTables in Excel UI

**CLI Usage**:

```powershell
# Create PivotTable slicer
excelcli slicer create-slicer --session <id> --pivot-table-name "SalesPivot" --field-name "Region" --destination-sheet "Dashboard"

# Set slicer filter
excelcli slicer set-slicer-selection --session <id> --slicer-name "RegionSlicer" --selected-items "[`"West`",`"East`"]"

# Clear slicer filter (show all)
excelcli slicer set-slicer-selection --session <id> --slicer-name "RegionSlicer" --selected-items "[]"

# Create Table slicer
excelcli slicer create-table-slicer --session <id> --table-name "SalesTable" --column-name "Category"

# List all slicers
excelcli slicer list-slicers --session <id>
excelcli slicer list-table-slicers --session <id>
```

````
