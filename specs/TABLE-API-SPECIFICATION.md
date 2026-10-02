# Excel Table (ListObject) API Specification

**Platform scope:** this Windows COM design/review is not a Mac action catalog.
All Table and column actions are unsupported in the experimental Apple Silicon
macOS beta; optional Office.js handlers remain candidates. See
[current feature guidance](../docs/features/CELLS-WORKBOOKS.md) and
[Mac limitations](MACOS-SUPPORT.md#not-supported-in-the-macos-beta).

> **Comprehensive specification for Excel Table operations - reviewing current implementation and future refactoring needs**

## Executive Summary

This specification reviews the **current TableCommands implementation** to determine:
1. What functionality already exists
2. What overlaps with RangeCommands
3. What should be refactored or removed
4. What's missing that should be added

### Key Questions to Answer

1. **Does TableCommands duplicate RangeCommands?** - Data operations on tables
2. **What's the proper division of responsibilities?** - Table structure vs data operations
3. **Should ReadDataAsync/AppendRowsAsync move to RangeCommands?** - Data operations
4. **What table-specific features are missing?** - Structured references, filters, slicers?

---

## Current TableCommands Implementation

### Interface Review (ITableCommands.cs)

**Lifecycle Operations:**
- ✅ `List` - List all tables in workbook
- ✅ `Create` - Create table from range with headers/style
- ✅ `Rename` - Rename table
- ✅ `Delete` - Delete table (convert back to range)
- ✅ `GetInfo` - Get detailed table information

**Structure Operations:**
- ✅ `Resize` - Resize table to new range
- ✅ `ToggleTotals` - Show/hide totals row
- ✅ `SetColumnTotal` - Set totals function for column
- ✅ `SetStyle` - Change table style

**Data Operations:** ⚠️ **POTENTIAL OVERLAP WITH RANGECOMMANDS**
- ✅ `ReadData` - Read table data
- ✅ `AppendRows` - Append rows to table

**Data Model Integration:**
- ✅ `AddToDataModelAsync` - Add table to Power Pivot

---

## Excel Table (ListObject) Capabilities

### What is an Excel Table?

Excel Tables (ListObject COM objects) are **structured ranges with metadata**:
- Named references (e.g., "SalesTable")
- Column headers with names
- Automatic expansion when data added
- Built-in filtering and sorting UI
- Table styles and formatting
- Totals row with aggregate functions
- Structured references in formulas (`[@ColumnName]`)
- Can be added to Data Model for relationships

### Excel COM API - ListObject Operations

#### 1. **Table Lifecycle**
```csharp
// Create table
dynamic listObjects = sheet.ListObjects;
dynamic table = listObjects.Add(
    SourceType: xlSrcRange,
    Source: sheet.Range["A1:D100"],
    XlListObjectHasHeaders: xlYes
);
table.Name = "SalesTable";

// Delete table (convert to range, preserve data)
table.Unlist();

// Delete table (remove everything)
table.Delete();
```

#### 2. **Table Properties**
```csharp
// Basic properties
string name = table.Name;
string range = table.Range.Address;
bool hasHeaders = table.ShowHeaders;
bool hasTotals = table.ShowTotals;
string style = table.TableStyle.Name;

// Row counts
int totalRows = table.Range.Rows.Count;  // Including header/totals
int dataRows = table.DataBodyRange?.Rows.Count ?? 0;  // Data only

// Column operations
int columnCount = table.ListColumns.Count;
dynamic column = table.ListColumns.Item(1);  // or by name
string columnName = column.Name;
```

#### 3. **Table Resize**
```csharp
// Resize table to new range
table.Resize(sheet.Range["A1:E200"]);
```

#### 4. **Totals Row**
```csharp
// Show/hide totals row
table.ShowTotals = true;

// Set totals function for column
dynamic column = table.ListColumns.Item("Amount");
column.TotalsCalculation = xlTotalsCalculationSum;  // Sum
column.TotalsCalculation = xlTotalsCalculationAverage;  // Average
column.TotalsCalculation = xlTotalsCalculationCount;  // Count
column.TotalsCalculation = xlTotalsCalculationMax;  // Max
column.TotalsCalculation = xlTotalsCalculationMin;  // Min
column.TotalsCalculation = xlTotalsCalculationStdDev;  // Std Dev
column.TotalsCalculation = xlTotalsCalculationVar;  // Variance
column.TotalsCalculation = xlTotalsCalculationCustom;  // Custom formula
```

#### 5. **Table Styles**
```csharp
// Built-in styles
table.TableStyle = workbook.TableStyles.Item("TableStyleMedium2");

// Or by name
table.TableStyle = "TableStyleLight9";
```

#### 6. **AutoFilter (Filtering)**
```csharp
// Tables automatically have AutoFilter
dynamic autoFilter = table.AutoFilter;

// Apply filter to column
autoFilter.Range.AutoFilter(
    Field: 2,  // Column index (1-based)
    Criteria1: "USA",
    Operator: xlFilterValues
);

// Clear filters
autoFilter.ShowAllData();

// Check if filtered
bool isFiltered = table.ShowAutoFilter;
```

#### 7. **Data Operations**
```csharp
// Read data (values only, no headers)
dynamic dataBodyRange = table.DataBodyRange;
object[,] values = dataBodyRange?.Value2;

// Read entire table (including headers)
object[,] allData = table.Range.Value2;

// Append row (table auto-expands)
dynamic newRow = table.ListRows.Add();
newRow.Range.Value2 = new object[,] { { val1, val2, val3 } };

// Insert row at position
dynamic insertedRow = table.ListRows.Add(Position: 5);
```

#### 8. **Data Model Integration**
```csharp
// Add to Data Model (Power Pivot)
table.TableObject = table;  // Make it a "proper" table
// Then use Power Pivot to add to model
```

---

## Overlap Analysis: TableCommands vs RangeCommands

### Current Overlap

| Operation | TableCommands | RangeCommands | Verdict |
|-----------|--------------|---------------|---------|
| **Read data** | `ReadDataAsync` | `GetValuesAsync` | ⚠️ OVERLAP - RangeCommands can read table ranges |
| **Write data** | `AppendRowsAsync` | `SetValuesAsync` | ⚠️ OVERLAP - RangeCommands can write to table ranges |
| **Clear data** | ❌ Not implemented | `ClearContentsAsync` | ✅ RangeCommands handles this |
| **Format cells** | ❌ Not implemented | `SetNumberFormatAsync`, `SetFontAsync`, etc. | ✅ RangeCommands handles this |

### Key Insight: Tables ARE Ranges with Metadata

Excel Tables are fundamentally **ranges with additional structure**:
- Underlying cells = regular range
- Table structure = metadata layer (headers, totals, style, filters)

**Proposed Division:**
- **TableCommands** = Table **structure and metadata** (lifecycle, totals, filters, styles)
- **RangeCommands** = **Data operations** on any range (including table data ranges)

---

## Proposed Refactoring Strategy

### Option 1: Remove Data Operations from TableCommands

**Remove from TableCommands:**
- ❌ `ReadDataAsync` → Use `RangeCommands.GetValuesAsync(batch, sheetName, "TableName[#Data]")`
- ❌ `AppendRowsAsync` → Use `RangeCommands.SetValuesAsync` + `ResizeAsync`

**Keep in TableCommands:**
- ✅ All lifecycle operations (List, Create, Rename, Delete, GetInfo)
- ✅ All structure operations (Resize, ToggleTotals, SetColumnTotal, SetStyle)
- ✅ Data Model integration (AddToDataModel)
- ✅ Filter operations (if added)

**Benefits:**
- Clear separation: Table structure vs data operations
- Users learn ONE API for data (RangeCommands)
- TableCommands focuses on table-specific features

**Challenges:**
- Users need to know table structured references (`TableName[#Data]`)
- Auto-expansion on append requires manual resize

### Option 2: Keep Data Operations but Delegate to RangeCommands Internally

**Keep current interface:**
- ✅ `ReadDataAsync` - Internally calls RangeCommands
- ✅ `AppendRowsAsync` - Internally calls RangeCommands + auto-resize

**Benefits:**
- User-friendly API (no need to know structured references)
- Auto-expansion handled automatically
- Backwards compatible

**Challenges:**
- Duplication of functionality
- Two ways to do the same thing

### Option 3: Hybrid Approach (RECOMMENDED)

**TableCommands focuses on table-specific operations:**
- ✅ Lifecycle: List, Create, Rename, Delete, GetInfo
- ✅ Structure: Resize, ToggleTotals, SetColumnTotal, SetStyle
- ✅ Table-specific data: `AppendRows` (auto-expansion feature)
- ✅ Filters: Apply, clear, get filter state
- ✅ Data Model: AddToDataModel
- ❌ **Remove**: `ReadData` - Use RangeCommands instead

**Rationale:**
- `AppendRows` has table-specific behavior (auto-expansion) - KEEP
- `ReadData` is just range read with no table-specific logic - REMOVE
- Filters are table-specific (AutoFilter object) - ADD
- Data operations (format, copy, etc.) - Use RangeCommands

---

## Missing Table Features

### 1. **Filter Operations** ⭐ HIGH PRIORITY
```csharp
// Apply filter to column
Task<OperationResult> ApplyFilterAsync(IExcelBatch batch, string tableName, string columnName, string criteria);

// Apply multiple criteria filter
Task<OperationResult> ApplyFilterAsync(IExcelBatch batch, string tableName, string columnName, List<string> values);

// Clear filters
Task<OperationResult> ClearFiltersAsync(IExcelBatch batch, string tableName);

// Get filter state
Task<TableFilterResult> GetFiltersAsync(IExcelBatch batch, string tableName);
```

**Excel COM:**
```csharp
dynamic autoFilter = table.AutoFilter;
autoFilter.Range.AutoFilter(Field: 2, Criteria1: "USA");
autoFilter.ShowAllData();  // Clear all filters
```

### 2. **Structured Reference Support** ⭐ MEDIUM PRIORITY
```csharp
// Get structured reference for table regions
Task<OperationResult> GetStructuredReferenceAsync(IExcelBatch batch, string tableName, TableRegion region);

public enum TableRegion
{
    All,        // TableName[#All]
    Data,       // TableName[#Data]
    Headers,    // TableName[#Headers]
    Totals,     // TableName[#Totals]
    ThisRow     // TableName[@]
}
```

### 3. **Column Operations** ⭐ MEDIUM PRIORITY
```csharp
// Add column to table
Task<OperationResult> AddColumnAsync(IExcelBatch batch, string tableName, string columnName, int? position = null);

// Remove column from table
Task<OperationResult> RemoveColumnAsync(IExcelBatch batch, string tableName, string columnName);

// Rename column
Task<OperationResult> RenameColumnAsync(IExcelBatch batch, string tableName, string oldName, string newName);
```

**Excel COM:**
```csharp
dynamic newColumn = table.ListColumns.Add(Position: 3);
newColumn.Name = "NewColumn";
table.ListColumns.Item("OldColumn").Delete();
```

### 4. **Sort Operations** ⭐ LOW PRIORITY (RangeCommands has Sort)
Tables can use standard Range.Sort(), so RangeCommands.SortAsync works on table ranges.

### 5. **Data Validation on Columns** ⭐ LOW PRIORITY (RangeCommands has Validation)
RangeCommands validation operations work on table column ranges.

### 6. **Slicers** ✅ IMPLEMENTED

Table slicers provide visual filtering controls for Excel Tables:

```csharp
// Create slicer for table column
SlicerResult CreateTableSlicer(IExcelBatch batch, string tableName, 
    string columnName, string slicerName, string destinationSheet, string position);

// List all Table slicers (optionally filter by table name)
SlicerListResult ListTableSlicers(IExcelBatch batch, string? tableName = null);

// Set slicer selection (filter values)
SlicerResult SetTableSlicerSelection(IExcelBatch batch, string slicerName, 
    List<string> selectedItems, bool clearFirst = true);

// Delete a Table slicer
OperationResult DeleteTableSlicer(IExcelBatch batch, string slicerName);
```

**Key Implementation Details:**
- Table slicers use `SlicerCaches.Add(table, columnName)` (deprecated but required for non-OLAP sources)
- Detection uses `SlicerCache.List` boolean property (returns `true` for Table slicers)
- Cannot use `SlicerCaches.Add2()` for Table slicers - it only supports PivotTable sources
- Exposed via `slicer` MCP tool with `create-table-slicer`, `list-table-slicers`, etc.

---

## Recommended TableCommands Refactoring

### Phase 1: Remove Duplication (THIS PHASE)

**Remove from TableCommands:**
1. ❌ `ReadDataAsync` - Users should use `RangeCommands.GetValuesAsync(batch, sheetName, "TableName[#Data]")`
   - Document migration: "Use RangeCommands to read table data"
   - Provide examples in docs

**Keep in TableCommands:**
2. ✅ `AppendRowsAsync` - Table-specific auto-expansion behavior
   - This is unique to tables (auto-resize when data added)
   - Cannot be easily replicated with RangeCommands alone

**Update Documentation:**
3. Document that RangeCommands works with table ranges
4. Provide examples of table structured references

### Phase 2: Add Missing Features (FUTURE)

**Filter Operations:**
1. `ApplyFilterAsync` - Apply filter to column
2. `ClearFiltersAsync` - Clear all filters
3. `GetFiltersAsync` - Get current filter state

**Column Operations:**
4. `AddColumnAsync` - Add column to table
5. `RemoveColumnAsync` - Remove column
6. `RenameColumnAsync` - Rename column

**Structured References:**
7. `GetStructuredReferenceAsync` - Get range address for table regions

---

## Implementation Details

### Current TableCommands Methods Review

#### ✅ KEEP - Table Lifecycle
- `ListAsync` - List all tables
- `CreateAsync` - Create table from range
- `RenameAsync` - Rename table
- `DeleteAsync` - Delete table
- `GetInfoAsync` - Get table details

#### ✅ KEEP - Table Structure
- `ResizeAsync` - Resize table
- `ToggleTotalsAsync` - Show/hide totals row
- `SetColumnTotalAsync` - Set totals function
- `SetStyleAsync` - Change table style

#### ✅ KEEP - Table-Specific Data
- `AppendRowsAsync` - Append with auto-expansion

#### ✅ KEEP - Data Model
- `AddToDataModelAsync` - Add to Power Pivot

#### ❌ REMOVE - Data Operations (Use RangeCommands)
- `ReadDataAsync` - Duplicate of RangeCommands.GetValuesAsync

---

## Migration Guide for Users

### Before (TableCommands.ReadDataAsync)
```csharp
var result = await tableCommands.ReadDataAsync(batch, "SalesTable");
List<List<object?>> data = result.Data;
```

### After (RangeCommands.GetValuesAsync)
```csharp
// Option 1: Read data only (no headers)
var result = await rangeCommands.GetValuesAsync(batch, "Sales", "SalesTable[#Data]");
List<List<object?>> data = result.Values;

// Option 2: Read everything (headers + data)
var result = await rangeCommands.GetValuesAsync(batch, "Sales", "SalesTable[#All]");

// Option 3: If you don't know the sheet name
var tableInfo = await tableCommands.GetInfoAsync(batch, "SalesTable");
var result = await rangeCommands.GetValuesAsync(batch, tableInfo.SheetName, "SalesTable[#Data]");
```

### Table Structured References

Excel Tables support structured references:
- `TableName[#All]` - Entire table including headers and totals
- `TableName[#Data]` - Data rows only (no headers or totals)
- `TableName[#Headers]` - Header row only
- `TableName[#Totals]` - Totals row only
- `TableName[[ColumnName]]` - Specific column
- `TableName[@]` - Current row (in formulas)

---

## Summary

### Current State
- TableCommands has 12 operations
- 1 operation (`ReadDataAsync`) duplicates RangeCommands
- 1 operation (`AppendRowsAsync`) has table-specific behavior worth keeping
- Missing important table features: filters, column management

### Proposed Changes

**Phase 1 - Remove Duplication:**
1. ❌ Delete `ReadDataAsync` - Use RangeCommands instead
2. ✅ Keep `AppendRowsAsync` - Table-specific auto-expansion
3. 📝 Update documentation with migration guide and structured reference examples

**Phase 2 - Add Missing Features (Future):**
4. Add filter operations (Apply, Clear, Get)
5. Add column operations (Add, Remove, Rename)
6. Add structured reference helper

### Architecture Principle

**TableCommands** = Table **structure and metadata**
- Lifecycle (create, rename, delete, list)
- Structure (resize, totals, styles, columns)
- Table-specific behaviors (auto-expansion on append, filters)
- Data Model integration
- **Table slicers** ✅ IMPLEMENTED

**RangeCommands** = **Data operations** on any range
- Read/write values and formulas
- Formatting and styling
- Copy, clear, insert, delete
- Works on table ranges via structured references

This maintains clear separation of concerns and prevents duplication!

---

## Open Questions for Review

1. **Should we remove `ReadDataAsync` in Phase 1?** 
   - Pro: Eliminates duplication, encourages unified API
   - Con: Breaking change, users need to learn structured references

2. **Should `AppendRowsAsync` accept 2D arrays instead of CSV?**
   - Pro: Consistent with RangeCommands (no CSV in Core/MCP)
   - Con: Breaking change

3. **Should we add filter operations in Phase 1 or Phase 2?**
   - Filters are common table operations
   - But adds scope to refactoring

4. ~~**Should we support table slicers?**~~ ✅ **DONE - Implemented in Issue #363**
   - Table slicers implemented via `ITableCommands` interface
   - Exposed via `slicer` MCP tool

**Next Step:** Review this specification and decide on Phase 1 scope before implementation!
