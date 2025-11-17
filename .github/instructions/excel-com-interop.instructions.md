---
applyTo: "src/ExcelMcp.Core/**/*.cs"
---

# Excel COM Interop Patterns

> **Essential patterns for Excel COM automation**

## Core Principles

1. **Use Late Binding** - `dynamic` types with `Type.GetTypeFromProgID()`
2. **1-Based Indexing** - Excel collections start at 1, not 0
4. **QueryTable Refresh REQUIRED** - `.Refresh(false)` synchronous for persistence
5. **NEVER use RefreshAll()** - Async/unreliable; use individual `connection.Refresh()` or `queryTable.Refresh(false)`

## Reference Resources

**NetOffice Library** - THE BEST source for ALL Excel COM Interop patterns:
- GitHub: https://github.com/NetOfficeFw/NetOffice
- **Use for ALL COM Interop work** - ranges, worksheets, workbooks, charts, PivotTables, Power Query, VBA, connections, everything
- NetOffice wraps Office COM APIs in strongly-typed C# - study their patterns for dynamic interop conversion
- Search NetOffice repository BEFORE implementing any Excel COM automation
- Particularly valuable for: PivotTables, OLAP CubeFields, Data Model operations, QueryTables, complex COM scenarios

## Resource Management

### ✅ ALWAYS use ExcelHelper.WithExcel()

```csharp
return ExcelHelper.WithExcel(filePath, save: false, (excel, workbook) =>
{
    dynamic? query = null;
    try {
        query = workbook.Queries.Item(1);
        // Use query...
    } finally {
        ExcelHelper.ReleaseComObject(ref query);
    }
    return 0;
});
```

**Handles:** Excel.Application creation, Workbook open/close, Excel.Quit(), COM cleanup, GC collection

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

**📚 Reference:** [Excel Object Model](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
