# Excel Range API Specification

> **Comprehensive range operations for ExcelMcp - replacing fragmented cell/sheet operations**

## Executive Summary

This specification defines a unified **Range API** that consolidates and replaces fragmented cell and partial sheet operations with a comprehensive, performance-optimized approach to working with Excel ranges.

### Key Design Decisions

1. **Single Cell = Range** - A single cell (e.g., "A1") is a 1x1 range; no separate "cell" API needed
2. **All Operations Use 2D Arrays** - Consistent interface whether operating on one cell or 10,000 cells
3. **COM-Backed Only** - Every operation uses native Excel COM (no data processing in server)
4. **CSV is CLI-Only** - Core/MCP use 2D arrays; CLI converts CSV for user convenience

### Goals

1. **Unify operations** - Single consistent API for single cells, ranges, entire columns/rows
2. **Performance** - Bulk operations using Excel COM's native 2D array support
3. **Replace fragmentation** - Eliminate duplication between CellCommands and SheetCommands
4. **Excel parity** - Support operations users expect from Excel UI
5. **Type safety** - Proper handling of values vs formulas vs formats

---

## Current State Analysis

### Existing Functionality (Fragmented)

**CellCommands** (Single cells only):
- ✅ GetValueAsync - Read single cell value
- ✅ SetValueAsync - Write single cell value  
- ✅ GetFormulaAsync - Read single cell formula
- ✅ SetFormulaAsync - Write single cell formula
- ❌ No formatting support
- ❌ No multi-cell support

**SheetCommands** (Worksheet-level only):
- ✅ ReadAsync - Read range data (values only)
- ✅ WriteAsync - Write CSV data to range
- ✅ ClearAsync - Clear range (values + formulas)
- ❌ No formula read/write
- ❌ No formatting support
- ❌ Inconsistent with Cell API

**HyperlinkCommands** (Exists, needs integration):
- ✅ AddHyperlinkAsync - Add hyperlink to cell/range
- ✅ RemoveHyperlinkAsync - Remove hyperlink
- ✅ ListHyperlinksAsync - List all hyperlinks in sheet
- ✅ GetHyperlinkAsync - Get hyperlink from cell
- ⚠️ Operates on ranges but separate API

**TableCommands** (Structured data, separate concern):
- ✅ Excel Table (ListObject) operations
- ✅ Styling, totals, data model integration
- ✅ Should remain separate (different abstraction level)

**NamedRangeCommands** (Named ranges, separate concern):
- ✅ Create, delete, update named range definitions
- ✅ List all named ranges
- ✅ Get/set single values (parameters)
- ✅ Should remain separate (named range lifecycle management)
- ⚠️ RangeCommands will ADD bulk read/write to named ranges (data operations)

### Problems with Current State

1. **API Confusion**: Cell vs Sheet operations overlap
2. **Performance**: Single-cell operations inefficient for bulk work
3. **Incomplete**: No formatting, no data validation, limited hyperlink integration
4. **Inconsistency**: Different patterns for similar operations

---

## Research: Excel Range Operations

### Core Operations (Must Have)

#### 1. **Value Operations**
```csharp
// Excel COM: Range.Value2 property (variant array)
range.Value2 = values;  // Bulk write
object[,] data = range.Value2;  // Bulk read
```

#### 2. **Formula Operations**  
```csharp
// Excel COM: Range.Formula property (string array)
range.Formula = formulas;  // Bulk write
object[,] formulas = range.Formula;  // Bulk read
```

#### 3. **Clearing**
```csharp
// Excel COM: Range.Clear() and variants
range.Clear();  // All content
range.ClearContents();  // Values/formulas only
range.ClearFormats();  // Formatting only
range.ClearComments();  // Comments only
range.ClearHyperlinks();  // Hyperlinks only
```

#### 4. **Copy/Paste**
```csharp
// Excel COM: Range.Copy() and Range.PasteSpecial()
sourceRange.Copy(destinationRange);  // Copy all
destinationRange.PasteSpecial(xlPasteValues);  // Paste values only
destinationRange.PasteSpecial(xlPasteFormulas);  // Paste formulas
destinationRange.PasteSpecial(xlPasteFormats);  // Paste formats
```

### Formatting Operations (Should Have)

#### 5. **Number Formatting**
```csharp
// Excel COM: Range.NumberFormat property
range.NumberFormat = "#,##0.00";  // Currency
range.NumberFormat = "0.00%";  // Percentage  
range.NumberFormat = "m/d/yyyy";  // Date
range.NumberFormat = "@";  // Text
```

#### 6. **Font Formatting**
```csharp
// Excel COM: Range.Font object
range.Font.Name = "Arial";
range.Font.Size = 12;
range.Font.Bold = true;
range.Font.Italic = true;
range.Font.Color = RGB(255, 0, 0);  // Red
```

#### 7. **Cell Formatting**
```csharp
// Excel COM: Range.Interior (background)
range.Interior.Color = RGB(255, 255, 0);  // Yellow background
range.Interior.Pattern = xlSolid;

// Range.Borders (borders)
range.Borders.LineStyle = xlContinuous;
range.Borders.Weight = xlMedium;
range.Borders.Color = RGB(0, 0, 0);
```

#### 8. **Alignment**
```csharp
// Excel COM: Range alignment properties
range.HorizontalAlignment = xlCenter;
range.VerticalAlignment = xlCenter;
range.WrapText = true;
range.Orientation = 45;  // Rotate text
```

### Smart Range Operations (Native Excel COM)

#### 9. **UsedRange**
```csharp
// Excel COM: Worksheet.UsedRange property
dynamic usedRange = sheet.UsedRange;
string address = usedRange.Address;  // e.g., "$A$1:$D$100"
object[,] values = usedRange.Value2;  // All non-empty data
```

#### 10. **CurrentRegion**
```csharp
// Excel COM: Range.CurrentRegion property
dynamic region = range.CurrentRegion;
string address = region.Address;  // Contiguous block around cell
object[,] values = region.Value2;
```

#### 11. **Range Properties**
```csharp
// Excel COM: Range information properties
string address = range.Address;  // Absolute address "$A$1:$D$10"
int rowCount = range.Rows.Count;  // Number of rows
int columnCount = range.Columns.Count;  // Number of columns
string numberFormat = range.NumberFormat;  // Format code
```

#### 12. **Named Ranges**
```csharp
// Excel COM: Workbook.Names collection
dynamic namedRange = workbook.Names.Item("SalesData").RefersToRange;
object[,] values = namedRange.Value2;  // Read from named range
namedRange.Value2 = newValues;  // Write to named range
```

#### 13. **Insert/Delete**
```csharp
// Excel COM: Range.Insert() and Range.Delete()
range.Insert(xlShiftDown);  // Insert cells, shift down
range.Insert(xlShiftToRight);  // Insert cells, shift right
range.Delete(xlShiftUp);  // Delete cells, shift up
range.Delete(xlShiftToLeft);  // Delete cells, shift left

// Entire rows/columns
range.EntireRow.Insert();  // Insert entire rows
range.EntireRow.Delete();  // Delete entire rows
range.EntireColumn.Insert();  // Insert entire columns
range.EntireColumn.Delete();  // Delete entire columns
```

#### 14. **Find/Replace**
```csharp
// Excel COM: Range.Find() and Range.Replace()
dynamic foundCell = range.Find(
    What: "searchText",
    LookIn: xlValues,  // or xlFormulas
    LookAt: xlWhole,   // or xlPart
    MatchCase: false
);

range.Replace(
    What: "oldText",
    Replacement: "newText",
    LookAt: xlPart,
    MatchCase: false
);
```

#### 15. **Sort**
```csharp
// Excel COM: Range.Sort()
range.Sort(
    Key1: range.Columns[1],  // First sort column
    Order1: xlAscending,
    Key2: range.Columns[2],  // Second sort column
    Order2: xlDescending,
    Header: xlYes  // Has headers
);
```

### Advanced Operations (Could Have)

#### 9. **Data Validation**
```csharp
// Excel COM: Range.Validation
range.Validation.Add(xlValidateList, xlValidAlertStop, xlBetween, "Item1,Item2,Item3");
range.Validation.Delete();
```

#### 10. **Conditional Formatting**
```csharp
// Excel COM: Range.FormatConditions
range.FormatConditions.Add(xlCellValue, xlGreater, "100");
range.FormatConditions(1).Interior.Color = RGB(255, 0, 0);
```

#### 11. **Merge/Unmerge**
```csharp
// Excel COM: Range.Merge/UnMerge
range.Merge();
range.UnMerge();
range.MergeCells;  // Property to check
```

#### 12. **Auto-Resize**
```csharp
// Excel COM: Range.AutoFit
range.Columns.AutoFit();  // Auto-size columns
range.Rows.AutoFit();  // Auto-size rows
```

---

## Proposed Range API Design

> **⚠️ IMPORTANT: All operations are backed by native Excel COM API**  
> **CSV Import/Export**: CLI-only feature (not in Core or MCP Server)  
> **Single Cell = Range**: A single cell (e.g., "A1") is treated as a 1x1 range

### Design Principles

1. **COM-Backed Only**: Every method uses native Excel COM operations
2. **No Data Processing**: Server doesn't transform data (transpose, statistics, etc.) - LLMs do that
3. **2D Arrays in Core**: Core uses `List<List<object?>>` (native C# representation)
4. **CSV in CLI Only**: CLI handles CSV ↔ 2D array conversion for user convenience
5. **JSON in MCP**: MCP Server serializes 2D arrays to JSON for protocol
6. **Single Cell = Range**: All operations work on ranges; single cells are 1x1 ranges (e.g., "A1" returns `[[value]]`)

### Phase 1: Core Operations (MVP)

#### IRangeCommands Interface

```csharp
public interface IRangeCommands
{
    // === VALUE OPERATIONS ===
    
    /// <summary>
    /// Gets values from a range as 2D array
    /// Single cell "A1" returns [[value]], range "A1:B2" returns [[v1,v2],[v3,v4]]
    /// </summary>
    Task<RangeValueResult> GetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Sets values in a range from 2D array
    /// Single cell "A1" accepts [[value]], range "A1:B2" accepts [[v1,v2],[v3,v4]]
    /// </summary>
    Task<OperationResult> SetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<object?>> values);
    
    // === FORMULA OPERATIONS ===
    
    /// <summary>
    /// Gets formulas from a range as 2D array (empty string if no formula)
    /// Single cell "A1" returns [["=SUM(B:B)"]], range "A1:B2" returns [[f1,f2],[f3,f4]]
    /// </summary>
    Task<RangeFormulaResult> GetFormulasAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Sets formulas in a range from 2D array
    /// </summary>
    Task<OperationResult> SetFormulasAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>> formulas);
    
    // === CLEAR OPERATIONS ===
    
    /// <summary>
    /// Clears all content (values, formulas, formats) from range
    /// </summary>
    Task<OperationResult> ClearAllAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Clears only values and formulas (preserves formatting)
    /// </summary>
    Task<OperationResult> ClearContentsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Clears only formatting (preserves values and formulas)
    /// </summary>
    Task<OperationResult> ClearFormatsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    // === COPY OPERATIONS ===
    
    /// <summary>
    /// Copies range to another location (all content)
    /// </summary>
    Task<OperationResult> CopyAsync(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange);
    
    /// <summary>
    /// Copies only values (no formulas or formatting)
    /// </summary>
    Task<OperationResult> CopyValuesAsync(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange);
    
    /// <summary>
    /// Copies only formulas (no values or formatting)
    /// </summary>
    Task<OperationResult> CopyFormulasAsync(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange);
    
    // === INSERT/DELETE OPERATIONS === (⭐ POWER USER ESSENTIAL)
    
    /// <summary>
    /// Inserts blank cells, shifting existing cells down or right
    /// </summary>
    Task<OperationResult> InsertCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress, InsertShiftDirection shift);
    
    /// <summary>
    /// Deletes cells, shifting remaining cells up or left
    /// </summary>
    Task<OperationResult> DeleteCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress, DeleteShiftDirection shift);
    
    /// <summary>
    /// Inserts entire rows above the range
    /// </summary>
    Task<OperationResult> InsertRowsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Deletes entire rows in the range
    /// </summary>
    Task<OperationResult> DeleteRowsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Inserts entire columns to the left of the range
    /// </summary>
    Task<OperationResult> InsertColumnsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Deletes entire columns in the range
    /// </summary>
    Task<OperationResult> DeleteColumnsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    // === FIND/REPLACE OPERATIONS === (⭐ POWER USER ESSENTIAL)
    
    /// <summary>
    /// Finds all cells matching criteria in range
    /// </summary>
    Task<RangeFindResult> FindAsync(IExcelBatch batch, string sheetName, string rangeAddress, string searchValue, FindOptions options);
    
    /// <summary>
    /// Replaces text/values in range
    /// </summary>
    Task<OperationResult> ReplaceAsync(IExcelBatch batch, string sheetName, string rangeAddress, string findValue, string replaceValue, ReplaceOptions options);
    
    // === SORT OPERATIONS === (⭐ POWER USER ESSENTIAL)
    
    /// <summary>
    /// Sorts range by one or more columns
    /// </summary>
    Task<OperationResult> SortAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<SortColumn> sortColumns, bool hasHeaders = true);
    
    // === NATIVE EXCEL COM OPERATIONS === (⭐ LLM/AI AGENT ESSENTIAL)
    
    /// <summary>
    /// Gets the used range (all non-empty cells) from worksheet
    /// Excel COM: Worksheet.UsedRange
    /// </summary>
    Task<RangeValueResult> GetUsedRangeAsync(IExcelBatch batch, string sheetName);
    
    /// <summary>
    /// Gets the current region (contiguous data block) around a cell
    /// Excel COM: Range.CurrentRegion
    /// </summary>
    Task<RangeValueResult> GetCurrentRegionAsync(IExcelBatch batch, string sheetName, string cellAddress);
    
    /// <summary>
    /// Gets range information (address, dimensions, number formats)
    /// Excel COM: Range.Address, Range.Rows.Count, Range.Columns.Count, Range.NumberFormat
    /// </summary>
    Task<RangeInfoResult> GetRangeInfoAsync(IExcelBatch batch, string sheetName, string rangeAddress);
}

// === SUPPORTING TYPES FOR PHASE 1 ===

public enum InsertShiftDirection { Down, Right }
public enum DeleteShiftDirection { Up, Left }

public class FindOptions
{
    public bool MatchCase { get; set; } = false;
    public bool MatchEntireCell { get; set; } = false;
    public bool SearchFormulas { get; set; } = true;
    public bool SearchValues { get; set; } = true;
    public bool SearchComments { get; set; } = false;
}

public class ReplaceOptions : FindOptions
{
    public bool ReplaceAll { get; set; } = true;
}

public class SortColumn
{
    public int ColumnIndex { get; set; }  // 1-based within range
    public bool Ascending { get; set; } = true;
}

public class RangeFindResult : OperationResult
{
    public List<RangeCell> MatchingCells { get; set; } = new();
}

public class RangeCell
{
    public string Address { get; set; } = string.Empty;  // e.g., "A5"
    public int Row { get; set; }
    public int Column { get; set; }
    public object? Value { get; set; }
}

public class RangeInfoResult : OperationResult
{
    public string Address { get; set; } = string.Empty;  // Absolute address from Excel COM
    public int RowCount { get; set; }                    // Excel COM: range.Rows.Count
    public int ColumnCount { get; set; }                 // Excel COM: range.Columns.Count
    public string? NumberFormat { get; set; }            // Excel COM: range.NumberFormat (first cell)
}
```

### Phase 2: Number Formatting

```csharp
public interface IRangeCommands
{
    // === NUMBER FORMAT OPERATIONS ===
    
    /// <summary>
    /// Gets number format codes from range
    /// </summary>
    Task<RangeFormatResult> GetNumberFormatsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Sets number format for entire range
    /// </summary>
    Task<OperationResult> SetNumberFormatAsync(IExcelBatch batch, string sheetName, string rangeAddress, string formatCode);
    
    /// <summary>
    /// Sets number formats from 2D array (cell-by-cell)
    /// </summary>
    Task<OperationResult> SetNumberFormatsAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>> formats);
}
```

### Phase 3: Visual Formatting

```csharp
public interface IRangeCommands
{
    // === FONT OPERATIONS ===
    
    /// <summary>
    /// Sets font properties for range
    /// </summary>
    Task<OperationResult> SetFontAsync(IExcelBatch batch, string sheetName, string rangeAddress, FontOptions font);
    
    // === CELL APPEARANCE ===
    
    /// <summary>
    /// Sets background color for range
    /// </summary>
    Task<OperationResult> SetBackgroundColorAsync(IExcelBatch batch, string sheetName, string rangeAddress, int color);
    
    /// <summary>
    /// Sets borders for range
    /// </summary>
    Task<OperationResult> SetBordersAsync(IExcelBatch batch, string sheetName, string rangeAddress, BorderOptions borders);
    
    /// <summary>
    /// Sets alignment for range
    /// </summary>
    Task<OperationResult> SetAlignmentAsync(IExcelBatch batch, string sheetName, string rangeAddress, AlignmentOptions alignment);
}

public class FontOptions
{
    public string? Name { get; set; }
    public int? Size { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public int? Color { get; set; }  // RGB color
}

public class BorderOptions
{
    public string LineStyle { get; set; } = "continuous";  // continuous, dashed, dotted, none
    public string Weight { get; set; } = "thin";  // thin, medium, thick
    public int? Color { get; set; }
}

public class AlignmentOptions
{
    public string? Horizontal { get; set; }  // left, center, right, justify
    public string? Vertical { get; set; }  // top, middle, bottom
    public bool? WrapText { get; set; }
}
```

### Phase 4: Advanced Features

```csharp
public interface IRangeCommands
{
    // === COMMENTS/NOTES === (⭐ POWER USER ESSENTIAL)
    
    /// <summary>
    /// Adds comment to a cell
    /// </summary>
    Task<OperationResult> AddCommentAsync(IExcelBatch batch, string sheetName, string cellAddress, string commentText, string? author = null);
    
    /// <summary>
    /// Gets all comments in range
    /// </summary>
    Task<RangeCommentsResult> GetCommentsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Deletes comment from a cell
    /// </summary>
    Task<OperationResult> DeleteCommentAsync(IExcelBatch batch, string sheetName, string cellAddress);
    
    // === PROTECTION === (⭐ POWER USER ESSENTIAL)
    
    /// <summary>
    /// Locks cells (prevents editing when sheet is protected)
    /// </summary>
    Task<OperationResult> LockCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Unlocks cells (allows editing even when sheet is protected)
    /// </summary>
    Task<OperationResult> UnlockCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Gets locked status of cells in range
    /// </summary>
    Task<RangeLockResult> GetLockedStatusAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    // === GROUPING/OUTLINE === (⭐ POWER USER ESSENTIAL)
    
    /// <summary>
    /// Groups rows (creates outline)
    /// </summary>
    Task<OperationResult> GroupRowsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Ungroups rows (removes outline)
    /// </summary>
    Task<OperationResult> UngroupRowsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Groups columns (creates outline)
    /// </summary>
    Task<OperationResult> GroupColumnsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Ungroups columns (removes outline)
    /// </summary>
    Task<OperationResult> UngroupColumnsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    // === DATA VALIDATION ===
    
    /// <summary>
    /// Adds data validation to range
    /// </summary>
    Task<OperationResult> AddValidationAsync(IExcelBatch batch, string sheetName, string rangeAddress, ValidationOptions validation);
    
    /// <summary>
    /// Removes data validation from range
    /// </summary>
    Task<OperationResult> RemoveValidationAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    // === MERGE OPERATIONS ===
    
    /// <summary>
    /// Merges cells in range
    /// </summary>
    Task<OperationResult> MergeCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Unmerges cells in range
    /// </summary>
    Task<OperationResult> UnmergeCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    // === AUTO-RESIZE ===
    
    /// <summary>
    /// Auto-fits column widths to content
    /// </summary>
    Task<OperationResult> AutoFitColumnsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Auto-fits row heights to content
    /// </summary>
    Task<OperationResult> AutoFitRowsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
}

// === SUPPORTING TYPES FOR PHASE 4 ===

public class RangeCommentsResult : OperationResult
{
    public List<CellComment> Comments { get; set; } = new();
}

public class CellComment
{
    public string CellAddress { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public string? Author { get; set; }
}

public class RangeLockResult : OperationResult
{
    public List<List<bool>> LockedStatus { get; set; } = new();  // 2D array matching range
}

public class ValidationOptions
{
    public string Type { get; set; } = "list";  // list, whole, decimal, date, time, text-length, custom
    public string Operator { get; set; } = "between";  // between, not-between, equal, not-equal, greater, less, greater-or-equal, less-or-equal
    public string? Formula1 { get; set; }  // First condition or list items
    public string? Formula2 { get; set; }  // Second condition (for between)
    public string? ErrorTitle { get; set; }
    public string? ErrorMessage { get; set; }
}
```

---

## Hyperlink Integration

**Decision**: Integrate hyperlinks directly into RangeCommands (DELETE HyperlinkCommands):

```csharp
public interface IRangeCommands
{
    // === HYPERLINK OPERATIONS ===
    
    /// <summary>
    /// Adds hyperlink to a single cell
    /// </summary>
    Task<OperationResult> AddHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress, string url, string? displayText = null, string? tooltip = null);
    
    /// <summary>
    /// Removes hyperlink from a single cell or all hyperlinks from a range
    /// </summary>
    Task<OperationResult> RemoveHyperlinkAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Lists all hyperlinks in a worksheet
    /// </summary>
    Task<RangeHyperlinkResult> ListHyperlinksAsync(IExcelBatch batch, string sheetName);
    
    /// <summary>
    /// Gets hyperlink from a specific cell
    /// </summary>
    Task<RangeHyperlinkResult> GetHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress);
}
```

**Rationale**: 
- Hyperlinks are just another property of a range/cell (like formulas or formatting)
- No need for separate command class
- Simpler API for users
- Consistent with Range-centric design

---

## Named Ranges Integration

### Unified Range Address Design

**Key Insight**: Named ranges are just **aliases for cell addresses**. RangeCommands should accept both seamlessly.

### Division of Responsibilities

**NamedRangeCommands** (Existing - KEEP):
- **Define** named ranges: `CreateAsync("SalesData", "Sheet1!A1:D100")`
- **Manage** named ranges: `UpdateAsync`, `DeleteAsync`, `ListAsync`
- **Get/Set single value**: `GetAsync`, `SetAsync` (treats named range as parameter/scalar)
- **Purpose**: Named range lifecycle and metadata management

**RangeCommands** (New - Unified Approach):
- **Accepts BOTH** `"Sheet1!A1:D100"` AND `"SalesData"` in rangeAddress parameter
- **Automatic resolution**: If rangeAddress is a named range, resolve to actual address internally
- **No separate methods needed**: `GetValuesAsync` works for both regular ranges and named ranges

### How It Works

```csharp
// STEP 1: Define the named range (NamedRangeCommands)
await NamedRangeCommands.CreateAsync(batch, "SalesData", "Sheet1!A1:D100");

// STEP 2: Write data - rangeAddress accepts BOTH formats
await rangeCommands.SetValuesAsync(batch, "", "SalesData", salesData);  // Named range
await rangeCommands.SetValuesAsync(batch, "Sheet1", "A1:D100", salesData);  // Regular range
// Both work identically!

// STEP 3: Read data - rangeAddress accepts BOTH formats
var result1 = await rangeCommands.GetValuesAsync(batch, "", "SalesData");  // Named range
var result2 = await rangeCommands.GetValuesAsync(batch, "Sheet1", "A1:D100");  // Regular range
// Both return same data!

// STEP 4: Update the named range reference (NamedRangeCommands)
await NamedRangeCommands.UpdateAsync(batch, "SalesData", "Sheet1!A1:D200");  // Expand range
```

### Implementation Strategy

```csharp
public async Task<RangeValueResult> GetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress)
{
    return await batch.ExecuteAsync(async (ctx, ct) =>
    {
        dynamic range;
        
        // Try to resolve as named range first
        if (string.IsNullOrEmpty(sheetName))
        {
            try
            {
                dynamic name = ctx.Book.Names.Item(rangeAddress);
                range = name.RefersToRange;  // Resolve named range to actual range
            }
            catch
            {
                throw new McpException($"Named range '{rangeAddress}' not found");
            }
        }
        else
        {
            // Regular sheet!range address
            dynamic sheet = ctx.Book.Worksheets.Item(sheetName);
            range = sheet.Range[rangeAddress];
        }
        
        // Rest of implementation identical for both paths
        object[,] values = range.Value2;
        // ...
    });
}
```

### API Comparison

| Operation | NamedRangeCommands | RangeCommands |
|-----------|------------------|---------------|
| **Create named range** | ✅ `CreateAsync` | ❌ |
| **Delete named range** | ✅ `DeleteAsync` | ❌ |
| **Update reference** | ✅ `UpdateAsync` | ❌ |
| **List all names** | ✅ `ListAsync` | ❌ |
| **Get single value** | ✅ `GetAsync` → scalar | ✅ `GetValuesAsync("", "Name")` → `[[value]]` |
| **Set single value** | ✅ `SetAsync` → scalar | ✅ `SetValuesAsync("", "Name", [[value]])` |
| **Get bulk data** | ❌ | ✅ `GetValuesAsync("", "Name")` → 2D array |
| **Set bulk data** | ❌ | ✅ `SetValuesAsync("", "Name", values)` |

### When to Use Which

**Use NamedRangeCommands when**:
- Defining/managing named range lifecycle (create, delete, update reference)
- Listing all named ranges in workbook
- Working with named ranges as single-value parameters (scalar get/set)

**Use RangeCommands when**:
- Reading/writing data (works with both named ranges and regular ranges)
- Don't need to know if it's a named range or not
- Want unified API for all range operations

### LLM/MCP Perspective

**As an LLM, I just want to read data - I don't care about the addressing scheme**:

```typescript
// I can use either - RangeCommands handles both
excel_range({ action: "get-values", sheetName: "", rangeAddress: "SalesData" })
excel_range({ action: "get-values", sheetName: "Sheet1", rangeAddress: "A1:D100" })

// No need for separate "get-named-range-values" action!
```

### Excel COM Mapping

```csharp
// Named range resolution (automatic in RangeCommands)
dynamic name = workbook.Names.Item("SalesData");
dynamic range = name.RefersToRange;  // Returns Range object
object[,] values = range.Value2;  // Same as regular range!

// Regular range (same final operation)
dynamic sheet = workbook.Worksheets.Item("Sheet1");
dynamic range = sheet.Range["A1:D100"];
object[,] values = range.Value2;  // Identical to named range!
```

**Key**: Both paths produce a `Range` COM object, so all operations are identical after resolution.

---

## Migration Strategy

### ⚠️ BREAKING CHANGES - No Backwards Compatibility Required

Since backwards compatibility is not required, we can make clean architectural decisions:

### Removal Plan

1. **CellCommands** → **DELETE ENTIRELY**
   - ❌ Remove ICellCommands interface
   - ❌ Remove CellCommands.cs implementation
   - ❌ Remove CLI CellCommands wrapper
   - ❌ Remove ExcelCellTool (MCP)
   - ❌ Remove all cell-* CLI commands
   - ❌ Remove excel_cell MCP tool
   - ✅ All functionality replaced by RangeCommands (single cell = range A1:A1)

2. **SheetCommands** → **REFACTOR & SIMPLIFY**
   - ❌ Remove ReadAsync (use RangeCommands.GetValuesAsync instead)
   - ❌ Remove WriteAsync (use RangeCommands.SetValuesAsync instead)
   - ❌ Remove ClearAsync (use RangeCommands.ClearContentsAsync instead)
   - ❌ Remove AppendAsync (use RangeCommands.SetValuesAsync with calculated range)
   - ✅ Keep sheet-level operations: List, Create, Rename, Delete, Copy (worksheet management)
   - ✅ SheetCommands becomes purely worksheet lifecycle management

3. **HyperlinkCommands** → **INTEGRATE INTO RANGECOMMANDS**
   - ❌ Remove IHyperlinkCommands interface
   - ❌ Remove HyperlinkCommands.cs implementation  
   - ❌ Remove CLI HyperlinkCommands wrapper
   - ❌ Remove all hyperlink-* CLI commands
   - ✅ Hyperlink operations become actions in RangeCommands
   - ✅ Hyperlink operations integrated into excel_range MCP tool

4. **Result Types** → **SIMPLIFY**
   - ❌ Remove CellValueResult (replaced by RangeValueResult)
   - ❌ Remove WorksheetDataResult (replaced by RangeValueResult)
   - ❌ Remove HyperlinkListResult (replaced by RangeHyperlinkResult)
   - ❌ Remove HyperlinkInfoResult (replaced by RangeHyperlinkResult)

### Clean Architecture Result

**Before** (Fragmented):
```
Commands/
├── CellCommands.cs          ← DELETE (4 methods)
├── SheetCommands.cs         ← SLIM DOWN (9 methods → 5 methods)
├── HyperlinkCommands.cs     ← DELETE (4 methods)
├── NamedRangeCommands.cs     ← KEEP (named ranges)
├── TableCommands.cs         ← KEEP (Excel tables)
└── ...
```

**After** (Unified):
```
Commands/
├── RangeCommands.cs         ← NEW (30+ methods, all range operations)
├── SheetCommands.cs         ← SIMPLIFIED (5 methods, worksheet lifecycle only)
├── NamedRangeCommands.cs     ← KEEP (named ranges)
├── TableCommands.cs         ← KEEP (Excel tables)
└── ...
```

### Implementation Strategy - MCP Server First, CLI Later

**Phase 1A** - Core Foundation (MCP Server Focus):
1. ✅ Create RangeValueResult and RangeFormulaResult models
2. ✅ Create IRangeCommands interface (all 40 methods defined)
3. ⬜ Implement RangeCommands.cs (all 40 methods with Excel COM)
4. ⬜ Create RangeCommandsTests.cs with comprehensive integration tests
5. ⬜ **MCP Server**: Create ExcelRangeTool with all actions
6. ⬜ **MCP Server**: Update ExcelTools.cs to route to new range tool
7. ⬜ **MCP Server**: Update server.json with range tool definition
8. ⬜ **MCP Server**: Test all range operations via MCP protocol
9. ⬜ **CLI**: ONLY update commands that break due to refactoring (sheet-read, sheet-write if affected)
10. ⬜ **DELETE**: Remove CellCommands from Core (breaks CLI cell-* commands - acceptable)
11. ⬜ **DELETE**: Remove HyperlinkCommands from Core (breaks CLI hyperlink-* commands - acceptable)
12. ⬜ **DELETE**: Remove excel_cell from MCP server (replaced by excel_range)

**Phase 1B** - CLI Implementation (After MCP Server Complete):
1. ⬜ Create CLI RangeCommands wrapper (ExcelMcp.CLI/Commands/RangeCommands.cs)
2. ⬜ Add range-* commands to Program.cs routing
3. ⬜ Add CLI tests for range commands
4. ⬜ Remove old CLI commands (cell-*, hyperlink-*, sheet data operations)
5. ⬜ Update README.md and installation guides

**Phase 2-4** - Future PRs (Number/Visual/Advanced Formatting):
- MCP Server first (add actions to excel_range tool)
- CLI later (add range-* subcommands)

### Breaking Changes Strategy

**MCP Server** (Phase 1A):
- ✅ excel_cell tool → REMOVED (replaced by excel_range)
- ✅ Cell operations in excel_worksheet → REMOVED (use excel_range)
- ✅ Hyperlink operations → MOVED to excel_range

**CLI** (Phase 1B):
- ⚠️ cell-* commands → REMOVED (replaced by range-* commands)
- ⚠️ hyperlink-* commands → REMOVED (replaced by range-* commands)
- ⚠️ sheet-read/write/clear/append → REFACTORED or REMOVED (use range-* commands)

**Core** (Phase 1A):
- ❌ ICellCommands / CellCommands → DELETED
- ❌ IHyperlinkCommands / HyperlinkCommands → DELETED
- ⚠️ ISheetCommands / SheetCommands → REFACTORED (lifecycle only)

---

## Implementation Order Details

**Phase 1A** (MCP Server - This PR):
- ✅ Create RangeValueResult and RangeFormulaResult models
- ✅ Create IRangeCommands interface (core operations + hyperlinks + native Excel COM)
- ⬜ Implement RangeCommands.cs (values, formulas, clear, copy, insert/delete, find/replace, sort, hyperlinks, Excel COM operations)
- ⬜ Create RangeCommandsTests.cs with comprehensive tests
- ⬜ **DELETE CellCommands** (interface, implementation, MCP tool, tests - CLI breaks temporarily)
- ⬜ **DELETE HyperlinkCommands** (interface, implementation, tests - CLI breaks temporarily)
- ⬜ **REFACTOR SheetCommands** (remove Read/Write/Clear/Append from Core, keep lifecycle)
- ⬜ **MCP**: Create ExcelRangeTool for MCP server (replacing excel_cell tool)
- ⬜ **MCP**: Update ExcelTools.cs routing
- ⬜ **MCP**: Update server.json configuration
- ⬜ **MCP**: Integration tests for excel_range tool
- ⬜ **CLI**: Minimal fixes for broken imports/references (don't add new range-* commands yet)
- ⬜ Update Core documentation and copilot instructions

**Phase 1B** (CLI - Follow-up PR):
- ⬜ Create CLI RangeCommands wrapper
- ⬜ Add range-* commands to Program.cs
- ⬜ Remove old CLI command implementations (cell-*, hyperlink-*)
- ⬜ Update CLI tests
- ⬜ Update README.md, INSTALLATION.md
- ⬜ Update CLI-specific copilot instructions

**Phase 2** (Future PR) - Number Formatting:
- Add number formatting operations to IRangeCommands
- Implement in RangeCommands.cs
- MCP: Add actions to excel_range tool
- CLI: Add range-format-* commands
- Tests for number formats
- Update documentation

**Phase 3** (Future PR) - Visual Formatting:
- Add visual formatting operations (fonts, colors, borders, alignment)
- Implement in RangeCommands.cs
- MCP: Add actions to excel_range tool
- CLI: Add range-style-* commands
- Tests for visual formatting
- Update documentation

**Phase 4** (Future PR) - Advanced Features:
- Add advanced features (validation, merge, auto-fit, comments, protection, grouping)
- Implement in RangeCommands.cs
- MCP: Add actions to excel_range tool
- CLI: Add range-advanced-* commands
- Tests for advanced features
- Update documentation

---

## Refactoring Existing Commands into Range API

### Analysis from LLM Perspective

**As an LLM using the MCP server**, I currently use different tools for similar operations:

#### Current Fragmentation

**Excel Worksheet Tool** (9 actions):
- `list` - List worksheets ✅ **KEEP** (metadata/lifecycle)
- `read` - Read range data → **MOVE TO** `excel_range.get-values`
- `write` - Write CSV to range → **MOVE TO** `excel_range.set-values` (CLI CSV conversion)
- `create` - Create worksheet ✅ **KEEP** (lifecycle)
- `rename` - Rename worksheet ✅ **KEEP** (lifecycle)
- `copy` - Copy worksheet ✅ **KEEP** (lifecycle)
- `delete` - Delete worksheet ✅ **KEEP** (lifecycle)
- `clear` - Clear range → **MOVE TO** `excel_range.clear`
- `append` - Append to range → **MOVE TO** `excel_range.append-values`

**Excel Cell Tool** (4 actions):
- `get-value` - Get single cell value → **REPLACE WITH** `excel_range.get-values` (1x1 range)
- `set-value` - Set single cell value → **REPLACE WITH** `excel_range.set-values` (1x1 range)
- `get-formula` - Get single cell formula → **REPLACE WITH** `excel_range.get-formulas` (1x1 range)
- `set-formula` - Set single cell formula → **REPLACE WITH** `excel_range.set-formulas` (1x1 range)

### Unified Design - LLM Perspective

**After refactoring, as an LLM I will have**:

**Excel Worksheet Tool** (5 actions) - Pure lifecycle management:
- `list` - List all worksheets
- `create` - Create new worksheet
- `rename` - Rename worksheet
- `copy` - Copy worksheet
- `delete` - Delete worksheet

**Excel Range Tool** (38+ actions) - All data operations:
- `get-values` - Read any range (single cell = 1x1, whole sheet = UsedRange)
- `set-values` - Write any range (replaces write, set-value)
- `append-values` - Append rows to range (replaces append)
- `clear` - Clear range (replaces clear, supports variants)
- `get-formulas`, `set-formulas` - Formula operations (replaces get-formula, set-formula)
- Plus 30+ more specialized range operations

### Impact Analysis

**Actions Deleted** (13 total):
- From `excel_worksheet`: `read`, `write`, `clear`, `append` (4 actions)
- From `excel_cell`: ALL 4 actions
- All hyperlink actions (if any): estimate 5 actions

**Actions Added** (38 new):
- All IRangeCommands methods become MCP actions

**Net Result**: More focused tools, clearer separation of concerns, unified interface.

### Code Refactoring Plan

#### Phase 1A: Core & MCP Server

**Delete**:
- `ICellCommands.cs`, `CellCommands.cs` (Core)
- `IHyperlinkCommands.cs`, `HyperlinkCommands.cs` (Core)
- `ExcelCellTool.cs` (MCP Server)
- Result types: `CellValueResult`, `HyperlinkListResult`, etc.

**Modify**:
- `ISheetCommands.cs` - Remove: `ReadAsync`, `WriteAsync`, `ClearAsync`, `AppendAsync`
- `ISheetCommands.cs` - Keep: `ListAsync`, `CreateAsync`, `RenameAsync`, `CopyAsync`, `DeleteAsync`
- `SheetCommands.cs` - Delete methods: `ReadAsync`, `WriteAsync`, `ClearAsync`, `AppendAsync`, `ParseCsv`
- `ExcelWorksheetTool.cs` - Remove actions: `read`, `write`, `clear`, `append`
- `ExcelWorksheetTool.cs` - Keep actions: `list`, `create`, `rename`, `copy`, `delete`

**Add**:
- `IRangeCommands.cs` (38 methods) - NEW
- `RangeCommands.cs` (implementation) - NEW
- `ExcelRangeTool.cs` (38 actions) - NEW
- Result types: `RangeValueResult`, `RangeFormulaResult`, etc. - MOSTLY DONE

#### Phase 1B: CLI

**Delete**:
- `CLI/Commands/CellCommands.cs`
- CLI commands: `cell-get-value`, `cell-set-value`, `cell-get-formula`, `cell-set-formula`

**Modify**:
- `CLI/Commands/SheetCommands.cs` - Remove: `Read`, `Write`, `Clear`, `Append`
- `CLI/Commands/SheetCommands.cs` - Keep: `List`, `Create`, `Rename`, `Copy`, `Delete`
- `CLI/Program.cs` - Remove routing for deleted commands

**Add**:
- `CLI/Commands/RangeCommands.cs` (wraps Core with CSV conversion)
- CLI commands: `range-get-values`, `range-set-values`, `range-append-values`, `range-clear`, `range-get-formulas`, `range-set-formulas`, etc.

### Migration Path for Users

**Before** (fragmented):
```typescript
// Read data - uses worksheet tool
excel_worksheet({ action: "read", excelPath: "data.xlsx", sheetName: "Sales", range: "A1:D100" })

// Get single cell - uses cell tool
excel_cell({ action: "get-value", excelPath: "data.xlsx", sheetName: "Sales", cell: "A1" })

// Clear range - uses worksheet tool
excel_worksheet({ action: "clear", excelPath: "data.xlsx", sheetName: "Sales", range: "A1:D100" })
```

**After** (unified):
```typescript
// Read data - uses range tool
excel_range({ action: "get-values", excelPath: "data.xlsx", sheetName: "Sales", rangeAddress: "A1:D100" })

// Get single cell - uses range tool (1x1 range)
excel_range({ action: "get-values", excelPath: "data.xlsx", sheetName: "Sales", rangeAddress: "A1" })
// Returns: { values: [[value]] }

// Clear range - uses range tool
excel_range({ action: "clear", excelPath: "data.xlsx", sheetName: "Sales", rangeAddress: "A1:D100" })

// Bonus: Read entire sheet
excel_range({ action: "get-used-range", excelPath: "data.xlsx", sheetName: "Sales" })
```

**Key LLM Benefits**:
1. **Single tool for all data operations** - No more guessing which tool to use
2. **Consistent interface** - All actions use `rangeAddress` parameter
3. **More powerful** - Access to UsedRange, CurrentRegion, Find, Sort, etc.
4. **Named ranges work transparently** - No separate methods needed

---

## Usage Examples - Single Cell vs Range

### Single Cell Operations (1x1 Range)

```csharp
// Get single cell value - returns 2D array with 1 row, 1 column
var result = await rangeCommands.GetValuesAsync(batch, "Sheet1", "A1");
// result.Values = [[100]]

// Set single cell value - accepts 2D array with 1 row, 1 column
await rangeCommands.SetValuesAsync(batch, "Sheet1", "A1", [[100]]);

// Get single cell formula
var formulaResult = await rangeCommands.GetFormulasAsync(batch, "Sheet1", "C5");
// formulaResult.Formulas = [["=SUM(A1:A10)"]]

// Set single cell formula
await rangeCommands.SetFormulasAsync(batch, "Sheet1", "C5", [["=SUM(A1:A10)"]]);
```

### Multi-Cell Range Operations

```csharp
// Get range values - returns 2D array
var result = await rangeCommands.GetValuesAsync(batch, "Sheet1", "A1:C3");
// result.Values = [
//   [1, 2, 3],
//   [4, 5, 6],
//   [7, 8, 9]
// ]

// Set range values
await rangeCommands.SetValuesAsync(batch, "Sheet1", "A1:C3", [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
]);
```

### MCP JSON Examples

```json
// Single cell get-values
{
  "action": "get-values",
  "sheetName": "Sheet1",
  "rangeAddress": "A1"
}
// Returns: { "values": [[100]] }

// Single cell set-values
{
  "action": "set-values",
  "sheetName": "Sheet1",
  "rangeAddress": "A1",
  "values": [[100]]
}

// Range get-values
{
  "action": "get-values",
  "sheetName": "Sheet1",
  "rangeAddress": "A1:C3"
}
// Returns: { "values": [[1,2,3],[4,5,6],[7,8,9]] }
```

### ⚠️ CRITICAL: Range Address Must Match Data Dimensions

**ALWAYS specify the full range address matching your data dimensions.**

```csharp
// ❌ WRONG: Single cell address with multi-cell data
await rangeCommands.SetValuesAsync(batch, "Sheet1", "A1", [
    ["Date", "Region", "Product", "Revenue"]  // 1x4 array
]);
// May only write "Date" to A1, losing other columns!

// ✅ CORRECT: Full range address
await rangeCommands.SetValuesAsync(batch, "Sheet1", "A1:D1", [
    ["Date", "Region", "Product", "Revenue"]
]);

// ❌ WRONG: Two separate calls for headers + data
await rangeCommands.SetValuesAsync(batch, "Sheet1", "A1", [["Date", "Region"]]);
await rangeCommands.SetValuesAsync(batch, "Sheet1", "A2", [[1, "North"], [2, "South"]]);

// ✅ CORRECT: Single call with full range
await rangeCommands.SetValuesAsync(batch, "Sheet1", "A1:B3", [
    ["Date", "Region"],     // Headers
    [1, "North"],           // Data row 1
    [2, "South"]            // Data row 2
]);
```

**Why:** Excel COM does not reliably auto-expand from single cell addresses. Specifying the exact range ensures all data is written correctly.

### CLI Examples

```bash
# Single cell - CLI may simplify to scalar for user convenience
excelcli range-get-values file.xlsx Sheet1 A1
# Output: 100 (CLI unpacks [[100]] to scalar)

# Range - CLI displays as table or JSON
excelcli range-get-values file.xlsx Sheet1 A1:C3
# Output: Table or JSON 2D array

# Single cell from CSV (CLI converts to [[value]])
echo "100" > value.csv
excelcli range-set-values file.xlsx Sheet1 A1 value.csv

# Range from CSV (CLI converts to 2D array)
excelcli range-set-values file.xlsx Sheet1 A1:C3 data.csv
```

---

## Result Types

### Already Created

```csharp
/// <summary>
/// Result for Excel range value operations
/// </summary>
public class RangeValueResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
    public List<List<object?>> Values { get; set; } = new();
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
}

/// <summary>
/// Result for Excel range formula operations
/// </summary>
public class RangeFormulaResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
    public List<List<string>> Formulas { get; set; } = new();
    public List<List<object?>> Values { get; set; } = new();
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
}
```

### Need to Add (Phase 2+)

```csharp
public class RangeFormatResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
    public List<List<string>> NumberFormats { get; set; } = new();
}

public class RangeHyperlinkResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
    public List<HyperlinkInfo> Hyperlinks { get; set; } = new();
}
```

---

## CLI Commands

> **⚠️ CSV Support is CLI-ONLY**  
> Core and MCP Server use `List<List<object?>>` (2D arrays).  
> CLI converts CSV ↔ 2D arrays for user convenience (testing, scripting).

### Phase 1 Commands (Replacing cell-*, hyperlink-*, and sheet data commands)

```bash
# === VALUE OPERATIONS (replaces cell-get-value, cell-set-value, sheet-read, sheet-write) ===
excelcli range-get-values <file.xlsx> <sheet> <range>                # Output: JSON or table
excelcli range-set-values <file.xlsx> <sheet> <range> <data.csv>     # CLI-ONLY: Reads CSV, converts to 2D array

# === FORMULA OPERATIONS (replaces cell-get-formula, cell-set-formula) ===
excelcli range-get-formulas <file.xlsx> <sheet> <range>              # Output: JSON or table
excelcli range-set-formulas <file.xlsx> <sheet> <range> <formulas.csv>  # CLI-ONLY: Reads CSV, converts to 2D array

# === CLEAR OPERATIONS (replaces sheet-clear) ===
excelcli range-clear-all <file.xlsx> <sheet> <range>
excelcli range-clear-contents <file.xlsx> <sheet> <range>
excelcli range-clear-formats <file.xlsx> <sheet> <range>

# === COPY OPERATIONS ===
excelcli range-copy <file.xlsx> <srcSheet> <srcRange> <tgtSheet> <tgtRange>
excelcli range-copy-values <file.xlsx> <srcSheet> <srcRange> <tgtSheet> <tgtRange>
excelcli range-copy-formulas <file.xlsx> <srcSheet> <srcRange> <tgtSheet> <tgtRange>

# === HYPERLINK OPERATIONS (replaces hyperlink-add, hyperlink-remove, hyperlink-list, hyperlink-get) ===
excelcli range-add-hyperlink <file.xlsx> <sheet> <cell> <url> [displayText] [tooltip]
excelcli range-remove-hyperlink <file.xlsx> <sheet> <range>
excelcli range-list-hyperlinks <file.xlsx> <sheet>
excelcli range-get-hyperlink <file.xlsx> <sheet> <cell>
```

### Removed Commands

```bash
# ❌ DELETED - Use range-* commands instead
cell-get-value
cell-set-value
cell-get-formula
cell-set-formula

# ❌ DELETED - Use range-* commands instead
hyperlink-add
hyperlink-remove
hyperlink-list
hyperlink-get

# ❌ DELETED - Use range-* commands instead
sheet-read          # Use range-get-values
sheet-write         # Use range-set-values
sheet-clear         # Use range-clear-*
sheet-append        # Use range-set-values with calculated range

# ✅ KEPT - Worksheet lifecycle management
sheet-list
sheet-create
sheet-rename
sheet-copy
sheet-delete
```

---

## MCP Tool: excel_range

> **⚠️ MCP Uses JSON, NOT CSV**  
> Parameters use JSON arrays (2D): `[[value1, value2], [value3, value4]]`  
> No CSV support in MCP Server.

### Phase 1 Actions (Replacing excel_cell tool)

```typescript
{
  "name": "excel_range",
  "description": "Comprehensive Excel range operations - values, formulas, hyperlinks, formatting, and more",
  "parameters": {
    "action": "string",
    "excelPath": "string",
    "sheetName": "string",
    "rangeAddress": "string",
    "values": "array<array<any>>",      // JSON 2D array, NOT CSV
    "formulas": "array<array<string>>", // JSON 2D array, NOT CSV
    // ... other parameters
  },
  "actions": [
    // Value operations (replaces excel_cell get-value, set-value)
    "get-values",      // Returns: { values: [[val1, val2], [val3, val4]] }
    "set-values",      // Input: { values: [[val1, val2], [val3, val4]] }
    
    // Formula operations (replaces excel_cell get-formula, set-formula)
    "get-formulas",    // Returns: { formulas: [["=A1+B1", "=SUM(A:A)"]] }
    "set-formulas",    // Input: { formulas: [["=A1+B1", "=SUM(A:A)"]] }
    
    // Clear operations
    "clear-all",
    "clear-contents",
    "clear-formats",
    
    // Copy operations
    "copy",
    "copy-values",
    "copy-formulas",
    
    // Hyperlink operations (replaces excel_hyperlink tool)
    "add-hyperlink",
    "remove-hyperlink",
    "list-hyperlinks",
    "get-hyperlink"
  ]
}
```

### Removed MCP Tools

```typescript
// ❌ DELETED - Replaced by excel_range
{
  "name": "excel_cell",  // All actions moved to excel_range
  "actions": ["get-value", "set-value", "get-formula", "set-formula"]
}

// ❌ DELETED - Replaced by excel_range  
{
  "name": "excel_hyperlink",  // All actions moved to excel_range
  "actions": ["add", "remove", "list", "get"]
}
```

### Modified MCP Tool: excel_worksheet

```typescript
{
  "name": "excel_worksheet",
  "description": "Worksheet lifecycle management - create, rename, copy, delete sheets",
  "actions": [
    "list",      // ✅ KEPT
    "create",    // ✅ KEPT
    "rename",    // ✅ KEPT
    "copy",      // ✅ KEPT
    "delete",    // ✅ KEPT
    // ❌ REMOVED: "read", "write", "clear", "append" (use excel_range instead)
  ]
}
```

---

## Testing Strategy

### Unit Tests (Fast, No Excel)
- Input validation
- Range address parsing
- Error handling logic

### Integration Tests (Requires Excel)
- Single-cell operations
- Multi-cell operations  
- Large ranges (performance)
- Edge cases (merged cells, formulas referencing other sheets)
- Error conditions (invalid ranges, protected sheets)

### Test Data Patterns
```csharp
// Small range (2x2)
var values = new List<List<object?>> {
    new() { "A1", "B1" },
    new() { "A2", "B2" }
};

// Large range (1000x10) for performance testing
var largeData = GenerateTestData(rows: 1000, cols: 10);

// Formulas with references
var formulas = new List<List<string>> {
    new() { "=A1+B1", "=SUM(A1:B1)" },
    new() { "=A2*2", "=AVERAGE(A:A)" }
};
```

---

## Performance Considerations

1. **Bulk Operations**: Use Excel's native 2D array support
   - Single COM call for range vs N calls for N cells
   - 100x-1000x faster for large ranges

2. **Batching**: All operations use IExcelBatch pattern
   - Multiple range operations in single Excel session
   - Automatic save batching

3. **Memory**: Large ranges handled efficiently
   - Excel COM handles memory
   - Stream CSV data for huge ranges

## Relationship to Excel Tables (ListObjects)

### Question: Can RangeCommands Manipulate Tables?

**Answer: YES, with important distinctions**

Excel Tables (ListObjects) are **structured ranges with metadata**. They exist in two layers:

1. **The underlying range** - Can be manipulated with RangeCommands
2. **The table structure** - Requires TableCommands for metadata operations

### What RangeCommands CAN Do with Tables

✅ **Read/Write Data** - Access the underlying range
```csharp
// Read all data from a table (including headers)
var data = await rangeCommands.GetValuesAsync(batch, "Sheet1", "A1:D100");

// Write to specific cells within a table
await rangeCommands.SetValuesAsync(batch, "Sheet1", "B5:C10", newValues);

// Update formulas in calculated columns
await rangeCommands.SetFormulasAsync(batch, "Sheet1", "E2:E100", formulas);
```

✅ **Format Table Cells** - Style the underlying range
```csharp
// Format number columns in a table
await rangeCommands.SetNumberFormatAsync(batch, "Sheet1", "C2:C100", "$#,##0.00");

// Apply conditional formatting to table data
await rangeCommands.SetBackgroundColorAsync(batch, "Sheet1", "D2:D100", RGB(255, 200, 200));
```

✅ **Clear Table Data** (but preserves structure)
```csharp
// Clear data rows (keeps headers and table structure)
await rangeCommands.ClearContentsAsync(batch, "Sheet1", "A2:D100");
```

### What ONLY TableCommands Can Do

❌ **Table Structure Operations** - Requires TableCommands
```csharp
// Create a table from a range
await tableCommands.CreateAsync(batch, "Sheet1", "SalesTable", "A1:D100", hasHeaders: true);

// Resize table (add/remove columns or rows)
await tableCommands.ResizeAsync(batch, "SalesTable", "A1:F150");

// Change table style
await tableCommands.SetStyleAsync(batch, "SalesTable", "TableStyleMedium2");

// Toggle totals row
await tableCommands.ToggleTotalsAsync(batch, "SalesTable", showTotals: true);

// Set column total functions (SUM, AVERAGE, etc.)
await tableCommands.SetColumnTotalAsync(batch, "SalesTable", "Amount", "SUM");

// Add table to Data Model
await tableCommands.AddToDataModelAsync(batch, "SalesTable");
```

### Best Practices for Table Manipulation

**Scenario 1: Updating Table Data**
```csharp
// ✅ GOOD - Use RangeCommands for data operations
await rangeCommands.SetValuesAsync(batch, "Sheet1", "B2:B50", updatedPrices);
```

**Scenario 2: Reformatting Table Columns**
```csharp
// ✅ GOOD - Use RangeCommands for formatting
await rangeCommands.SetNumberFormatAsync(batch, "Sheet1", "C2:C100", "0.00%");
```

**Scenario 3: Adding Rows to Table**
```csharp
// ⚠️ OPTION A - Use TableCommands for automatic table expansion
await tableCommands.AppendRowsAsync(batch, "SalesTable", csvData);

// ⚠️ OPTION B - Use RangeCommands if you know exact range
await rangeCommands.SetValuesAsync(batch, "Sheet1", "A101:D105", newRows);
// Then resize table to include new rows
await tableCommands.ResizeAsync(batch, "SalesTable", "A1:D105");
```

**Scenario 4: Creating Calculated Columns**
```csharp
// ✅ BEST - Use RangeCommands for formulas
await rangeCommands.SetFormulasAsync(batch, "Sheet1", "E2:E100", 
    new List<List<string>> { 
        new() { "=[@Amount]*[@Quantity]" }  // Table structured reference
    });
```

### Recommended Approach

**Keep Both APIs** - They serve different purposes:

1. **TableCommands** - Table lifecycle and structure
   - Create, rename, delete tables
   - Resize, change styles
   - Totals row management
   - Data Model integration
   - Table-specific operations (AppendRows with auto-expansion)

2. **RangeCommands** - Data and formatting
   - Read/write values and formulas
   - Format cells (numbers, fonts, colors, borders)
   - Clear data
   - Copy/paste operations
   - Works on ANY range (tables or not)

### Updated Architecture Decision

```
Commands/
├── RangeCommands.cs      ← Data & formatting (any range, including tables)
├── TableCommands.cs      ← Table structure & lifecycle (ListObject metadata)
├── SheetCommands.cs      ← Worksheet lifecycle
├── NamedRangeCommands.cs  ← Named ranges
└── ...
```

**Rationale:**
- **RangeCommands** = Low-level, works everywhere
- **TableCommands** = High-level, table-specific features
- **Complementary**, not redundant

### Example: Complete Table Workflow

```csharp
// 1. Create table structure (TableCommands)
await tableCommands.CreateAsync(batch, "Sales", "SalesTable", "A1:D1", 
    hasHeaders: true, tableStyle: "TableStyleMedium2");

// 2. Populate with data (RangeCommands)
await rangeCommands.SetValuesAsync(batch, "Sales", "A2:D100", salesData);

// 3. Format currency column (RangeCommands)
await rangeCommands.SetNumberFormatAsync(batch, "Sales", "D2:D100", "$#,##0.00");

// 4. Add calculated column (RangeCommands)
await rangeCommands.SetFormulasAsync(batch, "Sales", "E2:E100", profitFormulas);

// 5. Add totals row (TableCommands)
await tableCommands.ToggleTotalsAsync(batch, "SalesTable", true);
await tableCommands.SetColumnTotalAsync(batch, "SalesTable", "Amount", "SUM");

// 6. Read results (RangeCommands)
var results = await rangeCommands.GetValuesAsync(batch, "Sales", "A1:E101");
```

This demonstrates how both APIs work together seamlessly!

---

## Success Criteria - MCP Server First Approach

### Phase 1A - MCP Server Complete (THIS PR)

**Core Implementation**:
- [ ] All 40 Phase 1 operations implemented in RangeCommands.cs
- [ ] All operations tested with comprehensive integration tests
- [ ] CellCommands completely deleted (interface, implementation, MCP, tests)
- [ ] HyperlinkCommands completely deleted (interface, implementation, tests)
- [ ] SheetCommands refactored (removed Read/Write/Clear/Append, kept lifecycle operations)
- [ ] All old result types removed (CellValueResult, WorksheetDataResult, HyperlinkListResult, HyperlinkInfoResult)

**MCP Server**:
- [ ] ExcelRangeTool created with all 40 actions
- [ ] excel_cell tool deleted (replaced by excel_range)
- [ ] ExcelTools.cs routing updated
- [ ] server.json configuration updated
- [ ] MCP integration tests passing (all range actions work via protocol)
- [ ] MCP prompts updated for range operations

**CLI Minimal Changes**:
- [ ] Import errors fixed (references to deleted commands)
- [ ] Broken tests removed/disabled temporarily
- [ ] Build succeeds (CLI commands may be missing functionality temporarily)

**Documentation**:
- [ ] Copilot instructions updated (.github/instructions/)
- [ ] Core architecture documentation updated
- [ ] Breaking changes documented

### Phase 1B - CLI Complete (Follow-up PR)

**CLI Implementation**:
- [ ] CLI RangeCommands wrapper created
- [ ] range-* commands added to Program.cs
- [ ] Old CLI commands deleted (cell-*, hyperlink-*, sheet data operations)
- [ ] CLI tests updated and passing

**Documentation**:
- [ ] README.md updated (breaking changes, migration guide)
- [ ] INSTALLATION.md updated if needed
- [ ] CLI-specific copilot instructions updated

**Performance**:
- [ ] Performance benchmarks show 10x+ improvement for bulk operations (GetValues vs multiple GetValue calls)

### Overall Success

- [ ] MCP Server provides complete range automation (Phase 1A complete)
- [ ] CLI provides complete range automation (Phase 1B complete)
- [ ] All tests passing with 90%+ coverage
- [ ] No regression in existing functionality (Power Query, VBA, Tables, Parameters)
- [ ] Breaking changes clearly documented with migration examples

---

## Power User Assessment & Missing Functionality

### ⭐ Critical Operations Added to Phase 1

Based on Excel power user workflows, the following operations were **ADDED to Phase 1** as essential:

1. **Insert/Delete Operations** (6 methods)
   - `InsertCellsAsync`, `DeleteCellsAsync` (shift cells)
   - `InsertRowsAsync`, `DeleteRowsAsync` (entire rows)
   - `InsertColumnsAsync`, `DeleteColumnsAsync` (entire columns)
   - **Why Critical**: Data manipulation workflows (inserting rows, removing blanks, restructuring)

2. **Find/Replace Operations** (2 methods)
   - `FindAsync` (search with options: match case, whole cell, formulas vs values)
   - `ReplaceAsync` (bulk replacement with regex-like patterns)
   - **Why Critical**: Data cleaning, standardization, error correction

3. **Sort Operations** (1 method)
   - `SortAsync` (multi-column sort with ascending/descending)
   - **Why Critical**: Data analysis, report preparation, ranked lists

### 🔄 Operations Moved to Separate Command Class

Some operations are **intentionally excluded** from RangeCommands because they belong in separate, specialized command classes:

1. **AutoFilter** → Create `IFilterCommands` (separate from RangeCommands)
   - Filtering is complex enough to warrant dedicated commands
   - Operations: Apply filter, modify filter criteria, clear filter, get filter state
   - **Why Separate**: AutoFilter has state management (active/inactive), multiple filter types (values, top 10, custom, date filters), requires reading back filter state

2. **PivotTables** → Future `IPivotCommands` (not in current scope)
   - Pivot tables are their own abstraction layer
   - Operations: Create pivot, add fields, set aggregation, refresh
   - **Why Separate**: Complex object model distinct from ranges

3. **Charts** → Future `IChartCommands` (not in current scope)
   - Charts reference ranges but are separate objects
   - Operations: Create chart, set series, modify axes, apply style
   - **Why Separate**: Charts are worksheet objects, not range operations

### ❌ Operations Intentionally Excluded (Not Power User Workflows)

These operations are **NOT included** because they're rarely automated or better handled through Excel UI:

1. **Freeze Panes** → Worksheet-level operation (not range-specific)
   - Belongs in `ISheetCommands.FreezePanesAsync(sheetName, cellAddress)` if needed
   - Rarely automated, mostly interactive user preference

2. **Print Areas/Page Setup** → Worksheet-level operation
   - Complex configuration rarely automated via COM
   - Better handled through Excel UI or templates

3. **Sparklines** → Specialized visualization (low automation value)
   - Rarely automated programmatically
   - Excel UI provides better visual design tools

4. **Conditional Formatting** → Deferred to Phase 4 or separate
   - Complex rule engine (icon sets, color scales, data bars, formulas)
   - May warrant separate `IConditionalFormattingCommands` in future
   - **Decision Needed**: Include in Phase 4 or create dedicated commands?

### ✅ Final Phase 1 Scope (Revised)

**Core Data Operations**:
- ✅ Get/Set Values (2D arrays)
- ✅ Get/Set Formulas (2D arrays)
- ✅ Clear (all/contents/formats/comments)
- ✅ Copy (all/values/formulas/formats)
- ✅ Insert/Delete cells/rows/columns
- ✅ Find/Replace
- ✅ Sort
- ✅ Hyperlinks (add/remove/list/get)

**Native Excel COM Operations** (AI Agent Essential):
- ✅ GetUsedRangeAsync - `Worksheet.UsedRange`
- ✅ GetCurrentRegionAsync - `Range.CurrentRegion`
- ✅ GetRangeInfoAsync - `Range.Address`, `Range.Rows.Count`, `Range.Columns.Count`, `Range.NumberFormat`
- ✅ GetNamedRangeValuesAsync / SetNamedRangeValuesAsync - `Workbook.Names("name").RefersToRange`

**Phase 1 Result**: ~40 methods covering 95% of daily Excel power user data manipulation workflows + essential AI agent discovery operations.

### 🤔 Open Questions for Architect

1. **AutoFilter Complexity**:
   - Create separate `IFilterCommands` now or defer to Phase 2?
   - Recommendation: **Create now** - filtering is essential for data workflows

2. **Conditional Formatting**:
   - Include in Phase 4 RangeCommands or create `IConditionalFormattingCommands`?
   - Recommendation: **Defer to separate commands** - too complex for Range API

3. **Comments vs Notes**:
   - Excel has "Comments" (threaded) and "Notes" (legacy)
   - Which should we support?
   - Recommendation: **Support both** - `AddCommentAsync` (threaded) and `AddNoteAsync` (legacy)

4. **Protection Granularity**:
   - Should protection be range-level (lock/unlock cells) or worksheet-level?
   - Current spec: Range-level `LockCellsAsync` / `UnlockCellsAsync`
   - Also need: Worksheet-level `ProtectSheetAsync` / `UnprotectSheetAsync` in SheetCommands
   - Recommendation: **Both** - range sets locked property, worksheet enables protection

---

## Updated Implementation Plan - MCP Server First

### Phase 1A: Core Range + MCP Server (THIS PR)
- **Target**: 40 methods + complete MCP integration
- **Timeline**: 4-5 days
- **Focus**: MCP Server functionality complete, CLI minimal changes
- **Scope**:
  - ✅ Core implementation (RangeCommands.cs with 40 methods)
  - ✅ Integration tests (RangeCommandsTests.cs)
  - ✅ MCP Server tool (ExcelRangeTool with all actions)
  - ✅ Delete obsolete Core commands (CellCommands, HyperlinkCommands)
  - ✅ Refactor SheetCommands (lifecycle only)
  - ⚠️ CLI minimal fixes (import errors only, no new range-* commands)
  - ✅ Core documentation updates

### Phase 1B: CLI Implementation (Follow-up PR)
- **Target**: Complete CLI range-* commands
- **Timeline**: 2-3 days
- **Focus**: Full CLI support for range operations
- **Scope**:
  - ✅ CLI RangeCommands wrapper
  - ✅ Add range-* commands to Program.cs
  - ✅ Remove old CLI commands (cell-*, hyperlink-*)
  - ✅ CLI tests
  - ✅ Update README.md, INSTALLATION.md

### Phase 1.5: AutoFilter Commands (Separate PR)
- **Create**: `IFilterCommands` interface and implementation
- **Target**: 5-7 methods for AutoFilter workflows
- **Timeline**: 1-2 days
- **Scope**:
  - ✅ MCP Server first (filter actions in excel_worksheet or new tool)
  - ✅ CLI later (filter-* commands)

### Phase 2: Number Formatting (Future PR)
- **Timeline**: 1 day
- **MCP First**: Add actions to excel_range
- **CLI Later**: Add range-format-* commands

### Phase 3: Visual Formatting (Future PR)
- **Timeline**: 2 days
- **MCP First**: Add actions to excel_range
- **CLI Later**: Add range-style-* commands

### Phase 4: Advanced Features (Future PR)
- **Timeline**: 2 days
- **MCP First**: Add actions to excel_range
- **CLI Later**: Add range-advanced-* commands

**Total Phase 1 (MCP + CLI)**: 6-8 days
**Total All Phases**: ~12-15 days for comprehensive Range API (MCP priority)
