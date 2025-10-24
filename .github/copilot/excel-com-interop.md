# Excel COM Interop Patterns

> **Essential patterns for working with Excel COM automation**

## Core Principles

1. **Use Late Binding** - `dynamic` types with `Type.GetTypeFromProgID()`
2. **1-Based Indexing** - Excel collections start at 1, not 0
3. **Release COM Objects** - Use `Marshal.ReleaseComObject()` and `GC.Collect()`
4. **Never call RefreshAll()** - It hangs. Refresh individual connections instead.

---

## Resource Management Pattern

### ✅ ALWAYS use ExcelHelper.WithExcel()

```csharp
return ExcelHelper.WithExcel(filePath, save: false, (excel, workbook) =>
{
    // Your code here - lifecycle managed automatically
    dynamic sheets = workbook.Worksheets;
    
    // Release intermediate COM objects
    dynamic? query = null;
    try {
        query = workbook.Queries.Item(1);
        // Use query...
    } finally {
        ExcelHelper.ReleaseComObject(ref query);
    }
    
    return 0;  // Success
});
```

**Handles:**
- Excel.Application creation
- Workbook open/close
- Excel.Quit()
- COM cleanup
- GC collection

---

## Critical COM Issues

### 1. Excel Collections Are 1-Based

```csharp
// ❌ WRONG
var first = collection.Item(0);  // Throws error!

// ✅ CORRECT
var first = collection.Item(1);  // First item

for (int i = 1; i <= collection.Count; i++) {
    var item = collection.Item(i);
}
```

### 2. Named Range Format

```csharp
// ❌ WRONG - Missing = prefix
namesCollection.Add("Param", "Sheet1!A1");  // RefersToRange fails!

// ✅ CORRECT
string ref = reference.StartsWith("=") ? reference : $"={reference}";
namesCollection.Add("Param", ref);  // Now RefersToRange works
```

### 3. Power Query Loading

```csharp
// ❌ WRONG - Causes "Value does not fall within expected range"
listObjects.Add(...);  // DO NOT USE!

// ✅ CORRECT - Use QueryTables
string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
dynamic queryTable = sheet.QueryTables.Add(connectionString, sheet.Range["A1"], commandText);
queryTable.Refresh(false);
```

### 4. Excel Busy Handling

```csharp
catch (COMException ex) when (ex.HResult == -2147417851)
{
    // RPC_E_SERVERCALL_RETRYLATER - Excel is busy
    AnsiConsole.MarkupLine("[red]Excel is busy. Close dialogs and retry.[/]");
}
```

---

## Power Query Patterns

### Check If Query Is Connection-Only

```csharp
bool isConnectionOnly = true;
dynamic connections = workbook.Connections;

for (int i = 1; i <= connections.Count; i++) {
    dynamic conn = connections.Item(i);
    if (conn.Name == queryName || conn.Name == $"Query - {queryName}") {
        isConnectionOnly = false;
        break;
    }
}
```

### Refresh Query

```csharp
// ❌ NEVER use RefreshAll() - it hangs!
// workbook.RefreshAll();  // FORBIDDEN

// ✅ Refresh via connection
dynamic? targetConnection = FindConnection(workbook, queryName);
if (targetConnection != null) {
    targetConnection.Refresh();
}
```

---

## Worksheet Operations

### Read Data

```csharp
dynamic sheet = workbook.Worksheets.Item(sheetName);
dynamic range = sheet.Range[rangeAddress];  // "A1:D10"
object[,] values = range.Value2;  // 2D array, 1-based!

for (int row = 1; row <= values.GetLength(0); row++) {
    for (int col = 1; col <= values.GetLength(1); col++) {
        object cell = values[row, col];
    }
}
```

### Write Data

```csharp
object[,] data = new object[rows, cols];  // Populate data

dynamic targetRange = sheet.Range[startCell, endCell];
targetRange.Value2 = data;  // Bulk write
```

---

## Connection Type Discrepancy

**⚠️ Excel COM runtime types don't match official spec!**

```csharp
// Runtime reality (NOT spec):
if (connType == 3) {  // TEXT file connections
    dynamic textConn = conn.TextConnection;
}
else if (connType == 4) {  // WEB query connections
    dynamic webConn = conn.WebConnection;
}
```

---

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| Not releasing intermediate objects | Use `try/finally` + `ReleaseComObject()` |
| Using 0-based indexing | Excel is 1-based |
| Calling `RefreshAll()` | Refresh individual connections |
| Missing `=` in named ranges | Always prefix with `=` |
| Using `ListObjects.Add()` for PQ | Use `QueryTables.Add()` |

---

**📚 Reference:** Excel Object Model - https://docs.microsoft.com/en-us/office/vba/api/overview/excel
