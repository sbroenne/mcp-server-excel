---
applyTo: "src/**/*.cs"
---

# Architecture Patterns

> **Core patterns for ExcelMcp development**

## .NET Class Design (MANDATORY)

**Official Docs:** [Framework Design Guidelines](https://learn.microsoft.com/en-us/dotnet/standard/design-guidelines/), [Partial Classes](https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/partial-classes-and-methods)

### Key Rules

1. **One Public Class Per File** - Standard .NET practice (System.Text.Json, ASP.NET Core, EF Core)
2. **File Name = Class Name** - `RangeCommands.cs` contains `RangeCommands`
3. **Partial Classes for Large Implementations** - Split 15+ method classes by feature domain
4. **Descriptive Names** - No over-optimization (`RangeCommands` ✅, `Commands` ❌)
5. **Folder = Organization, Not Identity** - `Commands/Range/RangeCommands.cs`

### Partial Class Pattern

**When:** Class has 15+ methods, multiple feature domains, team collaboration

**Structure:**
```
Commands/Range/
    IRangeCommands.cs           # Interface
    RangeCommands.cs            # Partial (constructor, DI)
    RangeCommands.Values.cs     # Partial (Get/Set values)
    RangeCommands.Formulas.cs   # Partial (formulas)
    RangeHelpers.cs             # Separate helper class
```

**Benefits:** Git-friendly, team-friendly, ~100-200 lines per file, mirrors .NET Framework patterns

---

## Command Pattern

### Structure
```
Commands/
├── IPowerQueryCommands.cs    # Interface
├── PowerQueryCommands.cs     # Implementation
```

### Routing (Program.cs)
```csharp
return args[0] switch
{
    "pq-list" => powerQuery.List(args),
    "sheet-read" => sheet.Read(args),
    _ => ShowHelp()
};
```

---

## Resource Management Pattern

### ExcelHelper.WithExcel()

**✅ ALWAYS use - never manage Excel lifecycle manually**

```csharp
public int MyCommand(string[] args)
{
    return ExcelHelper.WithExcel(filePath, save: false, (excel, workbook) =>
    {
        dynamic? queries = null;
        try {
            queries = workbook.Queries;
            // Use queries...
            return 0;
        } finally {
            ExcelHelper.ReleaseComObject(ref queries);  // Release!
        }
    });
}
```

**Handles:** Excel.Application creation/destruction, Workbook open/close, COM cleanup, GC collection

---

## MCP Server Resource-Based Tools

**11 Focused Tools:**
1. `excel_batch` - Batch session management (start, commit, discard)
2. `excel_connection` - Data connections (OLEDB, ODBC, Text, Web)
3. `excel_datamodel` - Power Pivot / Data Model / DAX measures
4. `excel_file` - File operations (create, close, test)
5. `excel_namedrange` - Named ranges (parameters)
6. `excel_pivottable` - PivotTables for interactive analysis
7. `excel_powerquery` - Power Query M code management
8. `excel_range` - Range operations (values, formulas, hyperlinks)
9. `excel_table` - Excel Tables (ListObjects)
10. `excel_vba` - VBA macros
11. `excel_worksheet` - Worksheet lifecycle management

### Action-Based Routing
```csharp
[McpServerTool]
public async Task<string> ExcelPowerQuery(string action, ...)
{
    return action.ToLowerInvariant() switch
    {
        "list" => ListPowerQueries(...),
        "view" => ViewPowerQuery(...),
        _ => throw new McpException($"Unknown action: {action}")
    };
}
```

---

## DRY Shared Utilities

**ExcelHelper Methods:** `FindConnection()`, `FindQuery()`, `GetConnectionTypeName()`, `IsPowerQueryConnection()`, `CreateQueryTable()`, `SanitizeConnectionString()`

**Why:** Prevents 60+ lines of duplicate code per feature

---

## Security-First Patterns

```csharp
// Always sanitize before output
string safe = SanitizeConnectionString(connectionString);

// Defaults
SavePassword = false  // Never export credentials by default
```

---

## Performance Patterns

### Minimize Workbook Opens
```csharp
// ✅ GOOD - Single session
ExcelHelper.WithExcel(filePath, save, (e, wb) => {
    Operation1(wb); Operation2(wb); Operation3(wb);
    return 0;
});

// ❌ AVOID - Multiple sessions (slow)
```

### Bulk Operations
```csharp
// ✅ GOOD - Bulk read
object[,] values = range.Value2;

// ❌ AVOID - Cell-by-cell (slow COM calls)
```

---

## Key Principles

1. **WithExcel() for everything** - Never manual lifecycle
2. **Release intermediate objects** - Prevents Excel hanging
3. **Batch/Session for MCP** - Multiple operations in single session
4. **Resource-based tools** - 6 tools, not 33+ operations
5. **DRY utilities** - Share common patterns
6. **Security defaults** - Never expose credentials
7. **Bulk operations** - Minimize COM round-trips
