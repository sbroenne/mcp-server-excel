---
applyTo: "src/**/*.cs"
---

# Architecture Patterns

> **Core patterns for ExcelMcp development**

## .NET Class Design Guidelines (MANDATORY)

> **Follow official Microsoft .NET Framework Design Guidelines**

### Official Documentation
- [Framework Design Guidelines](https://learn.microsoft.com/en-us/dotnet/standard/design-guidelines/)
- [Names of Classes, Structs, and Interfaces](https://learn.microsoft.com/en-us/dotnet/standard/design-guidelines/names-of-classes-structs-and-interfaces)
- [Partial Classes and Methods](https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/partial-classes-and-methods)
- [C# Coding Conventions](https://learn.microsoft.com/en-us/dotnet/csharp/fundamentals/coding-style/coding-conventions)

### Key Rules

**1. One Public Class Per File**
```csharp
// ✅ CORRECT - Each public class in its own file
// File: RangeCommands.cs
public class RangeCommands : IRangeCommands { }

// File: RangeHelpers.cs
public static class RangeHelpers { }
```

**Why:** Standard practice in .NET projects (see System.Text.Json, ASP.NET Core, Entity Framework Core). Improves discoverability, reduces merge conflicts, makes navigation easier.

**2. File Name Matches Class Name**
```csharp
// ✅ CORRECT
// File: RangeCommands.cs
public class RangeCommands { }

// ❌ WRONG
// File: Commands.cs  (too generic!)
public class RangeCommands { }
```

**3. Partial Classes for Large Implementations** ([Official Guideline](https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/partial-classes-and-methods))
```csharp
// File: RangeCommands.cs (main/constructor)
public partial class RangeCommands : IRangeCommands
{
    public RangeCommands() { }
}

// File: RangeCommands.Values.cs (feature group)
public partial class RangeCommands
{
    public Task<RangeValueResult> GetValuesAsync(...) { }
    public Task<OperationResult> SetValuesAsync(...) { }
}

// File: RangeCommands.Formulas.cs (feature group)
public partial class RangeCommands
{
    public Task<RangeFormulaResult> GetFormulasAsync(...) { }
    public Task<OperationResult> SetFormulasAsync(...) { }
}
```

**4. Keep Descriptive Names - No Over-Optimization**
```csharp
// ✅ CORRECT - Clear, unambiguous
namespace Sbroenne.ExcelMcp.Core.Commands.Range;
public partial class RangeCommands : IRangeCommands { }

// ❌ WRONG - Too generic, causes ambiguity
namespace Sbroenne.ExcelMcp.Core.Commands.Range;
public partial class Commands : IRangeCommands { }  // What kind of commands?
```

**5. Folder Structure Provides Organization, Not Identity**
```
Commands/Range/              ← Organization
    IRangeCommands.cs        ← Identity: IRangeCommands
    RangeCommands.cs         ← Identity: RangeCommands
    RangeCommands.Values.cs  ← Identity: RangeCommands (partial)
    RangeHelpers.cs          ← Identity: RangeHelpers (separate class)
```

### Large Command Class Pattern

**When to split using partial classes:**
- Class has 15+ public methods
- Multiple feature domains (values, formulas, editing, search)
- Team collaboration on same class
- Want git-friendly organization

**How to organize partial files:**
```
Commands/Range/
    IRangeCommands.cs                  ← Interface (full contract)
    RangeCommands.cs                   ← Partial (constructor, DI, shared state)
    RangeCommands.Values.cs            ← Partial (Get/Set values)
    RangeCommands.Formulas.cs          ← Partial (Get/Set formulas)
    RangeCommands.Editing.cs           ← Partial (Clear, Copy, Insert, Delete)
    RangeCommands.Search.cs            ← Partial (Find, Replace, Sort)
    RangeCommands.Discovery.cs         ← Partial (UsedRange, CurrentRegion, RangeInfo)
    RangeCommands.Hyperlinks.cs        ← Partial (Hyperlink operations)
    RangeHelpers.cs                    ← Separate static helper class

Tests/Integration/Range/
    RangeCommandsTests.cs              ← Base test class (shared fixture)
    RangeCommandsTests.Values.cs       ← Test partial (mirrors implementation)
    RangeCommandsTests.Formulas.cs
    RangeCommandsTests.Editing.cs
    RangeCommandsTests.Search.cs
    RangeCommandsTests.Discovery.cs
    RangeCommandsTests.Hyperlinks.cs
    RangeTestHelpers.cs                ← Test utilities (separate class)
```

**Benefits:**
- ✅ Follows .NET Framework patterns (DbContext, Controllers, etc.)
- ✅ Clear file-to-feature mapping (~100-200 lines per file)
- ✅ Git-friendly (isolated changes, fewer merge conflicts)
- ✅ Team-friendly (multiple developers, different features)
- ✅ Discoverable (obvious where to find code)
- ✅ Testable (test structure mirrors implementation)

---

## Command Pattern

### Structure
```
Commands/
├── IPowerQueryCommands.cs    # Interface
├── PowerQueryCommands.cs     # Implementation
├── ISheetCommands.cs
├── SheetCommands.cs
```

### Routing (Program.cs)
```csharp
return args[0] switch
{
    "pq-list" => powerQuery.List(args),
    "pq-view" => powerQuery.View(args),
    "sheet-read" => sheet.Read(args),
    _ => ShowHelp()
};
```

---

## Resource Management Pattern

### ExcelHelper.WithExcel()

**✅ ALWAYS use this - never manage Excel lifecycle manually!**

```csharp
public int MyCommand(string[] args)
{
    return ExcelHelper.WithExcel(filePath, save: false, (excel, workbook) =>
    {
        dynamic? queries = null;
        try {
            queries = workbook.Queries;
            
            for (int i = 1; i <= queries.Count; i++) {
                dynamic? query = null;
                try {
                    query = queries.Item(i);
                    // Use query...
                } finally {
                    ExcelHelper.ReleaseComObject(ref query);  // Release!
                }
            }
            
            return 0;
        } finally {
            ExcelHelper.ReleaseComObject(ref queries);  // Release!
        }
    });
}
```

**Handles:**
- Excel.Application creation/destruction
- Workbook.Open()/Close()
- Excel.Quit()
- COM cleanup (`Marshal.ReleaseComObject()`)
- Garbage collection (optimized 2-cycle pattern)
- Proper null assignment

---

## Excel Instance Pooling (MCP Server Only)

### Purpose
Reuse Excel instances across operations to eliminate ~2-5 second startup overhead.

### Configuration
```csharp
// Program.cs - MCP Server startup
var pool = new ExcelInstancePool(
    idleTimeout: TimeSpan.FromSeconds(60), 
    maxInstances: 10
);
ExcelHelper.InstancePool = pool;  // Enable globally
```

### Benefits
- ✅ **~95% faster** for cached workbooks (2-5 sec → <100ms)
- ✅ **Conversational workflows** - Multiple operations in quick succession
- ✅ **Auto cleanup** - Idle instances disposed after 60 seconds
- ✅ **Thread-safe** - Concurrent requests handled correctly
- ✅ **Resource limits** - Max 10 Excel instances prevents exhaustion
- ✅ **Zero code changes** - Core commands automatically use pooling

### CLI Behavior
**No pooling** - CLI uses simple single-instance pattern for reliability.

### Capacity Management
```csharp
// When pool is full, operations timeout after 5 seconds
// LLM can free slots:
excel_file({ 
  action: "close-workbook", 
  excelPath: "path/to/file.xlsx" 
})
// Returns: "Workbook closed in pool. Instance slot freed."
```

### Critical Fix: Semaphore Race Condition

**Problem:** TOCTOU bug between `ContainsKey()`, semaphore acquisition, and `GetOrAdd()`.

**Solution:** Atomic lock around entire sequence:
```csharp
lock (_instanceCreationLock)  // CRITICAL
{
    bool isExistingInstance = _instances.ContainsKey(normalizedPath);
    
    if (!isExistingInstance) {
        _instanceSemaphore.Wait(TimeSpan.FromSeconds(5));
        semaphoreAcquired = true;
    }
    
    pooledInstance = _instances.GetOrAdd(normalizedPath, _ => CreatePooledInstance(filePath));
}
```

**Why:** Without lock, multiple threads can check `ContainsKey()` simultaneously, both acquire semaphore, but only one instance created → semaphore count mismatch.

---

## MCP Server Resource-Based Tools

### Structure (6 Focused Tools)

1. **`excel_file`** - Excel file operations (1 action)
2. **`excel_powerquery`** - Power Query M code (11 actions)
3. **`excel_worksheet`** - Worksheet operations (9 actions)
4. **`excel_parameter`** - Named ranges (5 actions)
5. **`excel_cell`** - Individual cells (4 actions)
6. **`excel_vba`** - VBA macros (6 actions)

### Action-Based Routing
```csharp
[McpServerTool]
public async Task<string> ExcelPowerQuery(
    string action,
    string excelPath,
    string? queryName = null,
    ...)
{
    return action.ToLowerInvariant() switch
    {
        "list" => ListPowerQueries(...),
        "view" => ViewPowerQuery(...),
        "import" => await ImportPowerQuery(...),
        _ => throw new McpException($"Unknown action: {action}")
    };
}
```

### Error Handling
```csharp
// ✅ Throw McpException - framework serializes
throw new McpException($"{action} failed for '{filePath}': {ex.Message}");

// ❌ Don't return JSON errors
return JsonSerializer.Serialize(new { error = "..." });  // WRONG
```

---

## DRY Shared Utilities

### ExcelHelper Utilities
- `FindConnection()` - Locate connection by name
- `FindQuery()` - Locate Power Query by name
- `GetConnectionTypeName()` - Type identification
- `IsPowerQueryConnection()` - Detection
- `CreateQueryTable()` - Standard query loading
- `RemoveConnections()` - Cleanup
- `SanitizeConnectionString()` - Security (password masking)

### Why This Matters
Prevents 60+ lines of duplicate code per feature and ensures consistent behavior.

---

## Security-First Patterns

### Sensitive Data Handling
```csharp
// Always sanitize before output
string safe = SanitizeConnectionString(connectionString);
Console.WriteLine(safe);  // Passwords masked
```

### Defaults
- `SavePassword = false` - Never export credentials by default
- Require explicit parameters for security-sensitive operations
- Clear warnings when affecting security settings

---

## Performance Patterns

### Minimize Workbook Opens
```csharp
// ✅ GOOD - Single session
ExcelHelper.WithExcel(filePath, save, (e, wb) => {
    Operation1(wb);
    Operation2(wb);
    Operation3(wb);
    return 0;
});

// ❌ AVOID - Multiple sessions
ExcelHelper.WithExcel(filePath, false, (e, wb) => Operation1(wb));
ExcelHelper.WithExcel(filePath, false, (e, wb) => Operation2(wb));
ExcelHelper.WithExcel(filePath, false, (e, wb) => Operation3(wb));
```

### Bulk Operations
```csharp
// ✅ GOOD - Bulk read
object[,] values = range.Value2;

// ❌ AVOID - Cell-by-cell (slow COM calls)
for (each cell) value = cell.Value2;
```

---

## Key Principles

1. **WithExcel() for everything** - Never manual lifecycle
2. **Release intermediate objects** - Prevents Excel hanging
3. **Pooling for MCP only** - CLI stays simple
4. **Resource-based tools** - 6 tools, not 33+ operations
5. **DRY utilities** - Share common patterns
6. **Security defaults** - Never expose credentials
7. **Bulk operations** - Minimize COM round-trips
