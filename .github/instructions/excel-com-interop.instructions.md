---
applyTo: "src/ExcelMcp.Core/**/*.cs"
---

# Excel COM Interop Patterns

> **Essential patterns for Excel COM automation**

## Core Principles

1. **Use Late Binding** - `dynamic` types with `Type.GetTypeFromProgID()`
2. **1-Based Indexing** - Excel collections start at 1, not 0
3. **Exception Propagation** - Never wrap in try-catch, let batch.Execute() handle exceptions (see Exception Propagation section)
4. **QueryTable Refresh REQUIRED** - `.Refresh(false)` synchronous for persistence
5. **NEVER use RefreshAll()** - Async/unreliable; use individual `connection.Refresh()` or `queryTable.Refresh(false)`

## Reference Resources

**NetOffice Library** - THE BEST source for ALL Excel COM Interop patterns:
- GitHub: https://github.com/NetOfficeFw/NetOffice
- **Use for ALL COM Interop work** - ranges, worksheets, workbooks, charts, PivotTables, Power Query, VBA, connections, everything
- NetOffice wraps Office COM APIs in strongly-typed C# - study their patterns for dynamic interop conversion
- Search NetOffice repository BEFORE implementing any Excel COM automation
- Particularly valuable for: PivotTables, OLAP CubeFields, Data Model operations, QueryTables, complex COM scenarios

## Exception Propagation Pattern (CRITICAL)

**Core Commands: NEVER wrap operations in try-catch blocks that return error results. Let exceptions propagate naturally.**

```csharp
// ❌ WRONG: Catching and wrapping exceptions
public async Task<OperationResult> CreateAsync(IExcelBatch batch, string name)
{
    try
    {
        return await batch.Execute((ctx, ct) => {
            var item = ctx.Create(name);
            return ValueTask.FromResult(new OperationResult { Success = true });
        });
    }
    catch (Exception ex)
    {
        // ❌ WRONG: Double-wrapping suppresses the exception
        return new OperationResult { Success = false, ErrorMessage = ex.Message };
    }
}

// ✅ CORRECT: Let batch.Execute() handle exceptions via TaskCompletionSource
public async Task<OperationResult> CreateAsync(IExcelBatch batch, string name)
{
    return await batch.Execute((ctx, ct) => {
        var item = ctx.Create(name);
        return ValueTask.FromResult(new OperationResult { Success = true });
    });
    // Exception flows to batch.Execute() → caught via TaskCompletionSource
    // → Returns OperationResult { Success = false, ErrorMessage }
}

// ✅ CORRECT: Finally blocks are the right place for COM resource cleanup
public async Task<OperationResult> ComplexAsync(IExcelBatch batch, string name)
{
    dynamic? temp = null;
    try
    {
        return await batch.Execute((ctx, ct) => {
            temp = ctx.CreateTemp(name);
            // ... operation ...
            return ValueTask.FromResult(new OperationResult { Success = true });
        });
    }
    finally
    {
        // ✅ Finally for resource cleanup, NOT catch for error handling
        if (temp != null)
        {
            ComUtilities.Release(ref temp!);
        }
    }
}
```

**Why This Pattern:**
- `batch.Execute()` ALREADY captures exceptions via `TaskCompletionSource` 
- Inner try-catch suppresses exceptions, causing double-wrapping and lost stack context
- Finally blocks work perfectly for COM resource cleanup (which must happen regardless of exception)
- Exception occurs at correct layer (batch), not suppressed at method level

**Safe Exception Handling (Keep these):**
- ✅ Loop continuations: `catch { continue; }` (safe, recovers loop)
- ✅ Optional property access: `catch { value = null; }` (safe, uses fallback)
- ✅ Specific error routing: `catch (COMException ex) when (ex.HResult == code) { ... }` (specific, not general)
- ✅ Finally blocks: Resource cleanup for COM objects (always needed)

**Pattern to Remove:**
- ❌ `catch (Exception ex) { return new Result { Success = false, ErrorMessage = ex.Message }; }`

**Architecture:**
```
Core Command (NO try-catch wrapping)
  └─> await batch.Execute()
      └─> TaskCompletionSource captures exception
          └─> Returns OperationResult { Success = false, ErrorMessage }
```

---

## Resource Management

### ✅ Unified Shutdown Pattern (Current Standard)

**All workbook close and Excel quit operations use `ExcelShutdownService` with resilient retry:**

```csharp
// In ExcelBatch, ExcelSession, FileCommands:
ExcelShutdownService.CloseAndQuit(workbook, excel, save: false, filePath, logger);
```

**Shutdown Order:**
1. **Optional Save** - If `save=true`, calls `workbook.Save()` explicitly before close
2. **Close Workbook** - Calls `workbook.Close(save)` (save param controls Excel's prompt behavior)
3. **Release Workbook** - Releases COM reference via `ComUtilities.Release()`
4. **Quit Excel** - Calls `excel.Quit()` with exponential backoff retry (6 attempts, 200ms base delay)
5. **Release Excel** - Releases COM reference via `ComUtilities.Release()`
6. **Automatic GC** - RCW finalizers handle final cleanup automatically (no forced GC needed per Microsoft guidance)

**Resilience Features:**
- Uses `Microsoft.Extensions.Resilience` retry pipeline
- **Outer timeout (30s)**: Overall cancellation for Excel.Quit() - catches hung Excel (modal dialogs, deadlocks)
- **Inner retry**: Exponential backoff (200ms base, 2x factor, 6 attempts) for transient COM busy errors
- Retries on: `RPC_E_SERVERCALL_RETRYLATER` (-2147417851), `RPC_E_CALL_REJECTED` (-2147418111)
- Structured logging for diagnostics (attempt number, HResult, elapsed time)
- Continues with COM cleanup even if Quit fails/times out
- **STA thread join (10s)**: Short verification timeout after quit succeeds/fails

**Save Semantics:**
```csharp
// Discard changes (default for disposal paths)
ExcelShutdownService.CloseAndQuit(workbook, excel, save: false, filePath, logger);

// Save before close (for explicit save operations)
ExcelShutdownService.CloseAndQuit(workbook, excel, save: true, filePath, logger);
```

**Why Unified Service:**
- Eliminates duplicated try/catch blocks across `ExcelBatch`, `ExcelSession`, `FileCommands`
- Consistent retry behavior for all Excel quit operations
- Centralized logging and diagnostics
- Handles edge cases: disconnected COM proxies, hung Excel, modal dialogs

**Timeout Architecture (Proper Layering):**
```
Overall Quit Timeout: 30 seconds (outer)
  └─> Resilient Retry: 6 attempts with exponential backoff (inner, ~6s max)
      └─> Individual Quit() calls
  └─> STA Thread Join: 10 seconds (verification only)
```
- **30s quit timeout**: Catches truly hung Excel (modal dialogs, deadlocks) via CancellationToken
- **6-attempt retry**: Handles transient COM busy states within the 30s window
- **10s thread join**: Quick verification that cleanup finished (not a primary timeout mechanism)

## COM Object Cleanup Pattern (CRITICAL)

**ALWAYS use try-finally for COM object cleanup. NEVER use catch blocks to swallow exceptions.**

### ❌ WRONG Patterns

```csharp
// WRONG #1: COM cleanup in try block (won't execute if exception occurs)
try
{
    dynamic pivotLayout = chart.PivotLayout;
    dynamic pivotTable = pivotLayout.PivotTable;
    name = pivotTable.Name?.ToString() ?? string.Empty;
    ComUtilities.Release(ref pivotTable!);  // ❌ Won't execute if exception above!
    ComUtilities.Release(ref pivotLayout!);
}
catch
{
    name = "(unknown)";  // ❌ Swallows exception, causes COM leak
}

// WRONG #2: Empty catch block (swallows exceptions silently)
try
{
    dynamic item = GetItem();
    // ... operations ...
    ComUtilities.Release(ref item!);
}
catch
{
    // ❌ Empty catch - swallows exception, no cleanup
}
```

### ✅ CORRECT Pattern

```csharp
// CORRECT: Finally block ensures cleanup regardless of exceptions
dynamic? pivotLayout = null;
dynamic? pivotTable = null;
try
{
    pivotLayout = chart.PivotLayout;
    pivotTable = pivotLayout.PivotTable;
    name = pivotTable.Name?.ToString() ?? string.Empty;
}
finally
{
    // ✅ ALWAYS executes - exception or no exception
    if (pivotTable != null) ComUtilities.Release(ref pivotTable!);
    if (pivotLayout != null) ComUtilities.Release(ref pivotLayout!);
}
// ✅ Exception propagates naturally to batch.Execute()
```

**Pattern Requirements:**
1. **Declare COM objects as `dynamic?` nullable** before try block
2. **Initialize to `null`**
3. **Acquire COM objects in try block**
4. **Release in finally block** with null checks
5. **NO catch blocks** unless specific exception handling required
6. **NEVER catch to set fallback values** - let exceptions propagate

**Why This Matters:**
- Finally blocks execute **regardless** of exceptions (try succeeds or fails)
- COM objects leak if Release() not reached before exception
- Swallowing exceptions with catch blocks hides real problems from batch.Execute()
- Empty catch blocks are code smell - remove them entirely
- Let exceptions propagate naturally to batch.Execute() for proper error handling

**See Also:**
- CRITICAL-RULES.md Rule 22 for complete requirements
- CRITICAL-RULES.md Rule 1b for exception propagation pattern

## Critical COM Issues

### 1. Excel Collections Are 1-Based
```csharp
// ❌ WRONG: collection.Item(0)  
// ✅ CORRECT: collection.Item(1)
for (int i = 1; i <= collection.Count; i++) { var item = collection.Item(i); }
```

### 2. Named Range Format
```csharp
// ❌ WRONG: namesCollection.Add("Param", "Sheet1!A1");  // Missing =
// ✅ CORRECT: namesCollection.Add("Param", "=Sheet1!A1");
string ref = reference.StartsWith("=") ? reference : $"={reference}";
```

### 3. Power Query Loading
```csharp
// ❌ WRONG: listObjects.Add(...)  // Causes "Value does not fall within expected range"
// ✅ CORRECT: Use QueryTables with synchronous refresh
string cs = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
dynamic qt = sheet.QueryTables.Add(cs, sheet.Range["A1"], commandText);
qt.Refresh(false);  // CRITICAL: false = synchronous, ensures persistence
```

### 4. QueryTable Persistence Pattern

**⚠️ RefreshAll() does NOT persist QueryTables!**

```csharp
// ❌ WRONG: workbook.RefreshAll(); workbook.Save();  // QueryTable lost on reopen
// ✅ CORRECT: queryTable.Refresh(false); workbook.Save();  // Persists properly
```

**Why:** `RefreshAll()` is async. Individual `qt.Refresh(false)` is synchronous and required for disk persistence.

### 5. Numeric Property Type Conversions

**⚠️ ALL Excel COM numeric properties return `double`, NOT `int`!**

```csharp
// ❌ WRONG: Implicit conversion fails at runtime
int orientation = field.Orientation;  // Runtime error: Cannot convert double to int
int position = field.Position;        // Runtime error: Cannot convert double to int
int function = field.Function;        // Runtime error: Cannot convert double to int

// ✅ CORRECT: Explicit conversion required
int orientation = Convert.ToInt32(field.Orientation);
int position = Convert.ToInt32(field.Position);
int comFunction = Convert.ToInt32(field.Function);
```

**Common Properties Affected:**
- `PivotField.Orientation` → `double` (not `XlPivotFieldOrientation` enum)
- `PivotField.Position` → `double` (not `int`)
- `PivotField.Function` → `double` (not `XlConsolidationFunction` enum)
- `Range.Row`, `Range.Column` → `double` (not `int`)
- Any numeric property from Excel COM → assume `double`

**Date Properties:**
```csharp
// RefreshDate can be DateTime OR double (OLE date)
private static DateTime? GetRefreshDateSafe(dynamic refreshDate)
{
    if (refreshDate == null) return null;
    if (refreshDate is DateTime dt) return dt;
    if (refreshDate is double dbl) return DateTime.FromOADate(dbl);
    return null;
}
```

**Why:** Excel COM uses `VARIANT` types internally, which represent numbers as `double`. C# `dynamic` binding preserves this type.

### 6. Excel Busy Handling
```csharp
catch (COMException ex) when (ex.HResult == -2147417851)
{
    // RPC_E_SERVERCALL_RETRYLATER - Excel is busy
}
```

## Common Patterns

### Read Data
```csharp
dynamic range = sheet.Range["A1:D10"];
object[,] values = range.Value2;  // 2D array, 1-based indexing
```

### Write Data
```csharp
object[,] data = new object[rows, cols];
dynamic range = sheet.Range[startCell, endCell];
range.Value2 = data;  // Bulk write
```

### Refresh Query
```csharp
// ❌ NEVER: workbook.RefreshAll();  // Hangs!
// ✅ CORRECT: targetConnection.Refresh();
```

## Connection Type Discrepancy

**⚠️ Excel COM runtime types don't match spec!**
```csharp
if (connType == 3 || connType == 4) {  // TEXT files report as type 4 (WEB)
    try { var conn = connection.TextConnection; }
    catch { var conn = connection.WebConnection; }
}
```

## Data Model (Power Pivot) API Limitations

**⚠️ KNOWN LIMITATION: Hidden columns, relationships, and measures cannot be detected via Excel COM API**

When objects are marked "Hidden from client tools" in Power Pivot, the Excel COM API provides no way to detect this or retrieve them.

**Affected Objects:**

| Object | Available Properties | Missing |
|--------|---------------------|---------|
| `ModelTableColumn` | Application, Creator, DataType, Name, Parent | **NO IsHidden** |
| `ModelRelationship` | Application, Creator, ForeignKeyColumn, ForeignKeyTable, PrimaryKeyColumn, PrimaryKeyTable, Active | **NO IsHidden** |
| `ModelMeasure` | Application, AssociatedTable, Creator, Description, FormatInformation, Formula, Name, Parent | **NO IsHidden** |

**Alternative APIs that were investigated and DO NOT WORK:**

| Approach | Why It Doesn't Work |
|----------|---------------------|
| TOM (Tabular Object Model) | Requires `Microsoft.AnalysisServices.Tabular` library which cannot connect to Excel's embedded Analysis Services engine |
| XMLA queries | Excel's embedded AS engine doesn't expose a queryable endpoint for external XMLA connections |
| CubeField.ShowInFieldList | Only applies to PivotTable field visibility, not underlying Data Model hidden status |

**Bottom Line:** If a column, relationship, or measure is hidden in the Data Model, it cannot be seen or listed through the Excel COM API. This is a fundamental limitation of Microsoft's Excel automation interface.

---

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| 0-based indexing | Excel is 1-based |
| `RefreshAll()` | Use individual refresh |
| Missing `=` in ranges | Always prefix with `=` |
| `ListObjects.Add()` for PQ | Use `QueryTables.Add()` |
| Not releasing objects | `try/finally` + `ReleaseComObject()` |
| `int x = field.Property` | Use `Convert.ToInt32()` for ALL numeric properties |
| Assuming enum types | Numeric properties return `double`, convert to enum |
| Using TOM/XMLA for Data Model | Not accessible from Excel COM - use only ModelTable/ModelTableColumn APIs |

**📚 Reference:** [Excel Object Model](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
