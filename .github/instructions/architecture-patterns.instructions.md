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

## TWO EQUAL ENTRY POINTS (CRITICAL)

**ExcelMcp has TWO first-class entry points: MCP Server AND CLI.** Both must have:
- **Feature parity**: Every action in MCP must exist in CLI and vice versa
- **Parameter parity**: Same parameters, same defaults, same validation
- **Behavior parity**: Same Core command, same result format

When adding or changing ANY feature, ALWAYS update BOTH entry points. See Rule 24 (Post-Change Sync).

```
MCP Server (MCP tools, JSON-RPC) ──► In-process ExcelMcpService ──► Core Commands ──► Excel COM
CLI (command-line args, console)  ──► CLI Daemon (named pipe) ─────► Core Commands ──► Excel COM
```

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

**See excel-com-interop.instructions.md** for complete WithExcel() pattern and COM object lifecycle management.

---

## Exception Propagation Pattern (CRITICAL)

**Core Commands: Let exceptions propagate naturally** - Do NOT suppress with catch blocks that return error results.

```csharp
// ❌ WRONG: Suppressing exception with catch block
public async Task<OperationResult> SomeAsync(IExcelBatch batch, string param)
{
    try
    {
        return await batch.Execute((ctx, ct) => {
            // ... operation ...
            return ValueTask.FromResult(new OperationResult { Success = true });
        });
    }
    catch (Exception ex)
    {
        // ❌ WRONG: Catches exception and returns error result
        return new OperationResult 
        { 
            Success = false, 
            ErrorMessage = ex.Message 
        };
    }
}

// ✅ CORRECT: Let exception propagate through batch.Execute()
public async Task<OperationResult> SomeAsync(IExcelBatch batch, string param)
{
    return await batch.Execute((ctx, ct) => {
        // ... operation ...
        return ValueTask.FromResult(new OperationResult { Success = true });
    });
    // Exception automatically caught by batch.Execute() via TaskCompletionSource
    // Returns OperationResult { Success = false, ErrorMessage } from batch layer
}

// ✅ CORRECT: Finally blocks still allowed for COM resource cleanup
public async Task<OperationResult> ComplexAsync(IExcelBatch batch, string param)
{
    dynamic? connection = null;
    try
    {
        return await batch.Execute((ctx, ct) => {
            connection = ctx.Book.Connections.Add(...);
            // ... operation ...
            return ValueTask.FromResult(new OperationResult { Success = true });
        });
    }
    finally
    {
        if (connection != null)
        {
            ComUtilities.Release(ref connection!);  // ✅ Cleanup in finally
        }
    }
}
```

**Why This Pattern:**
- `batch.Execute()` already captures exceptions via `TaskCompletionSource`
- Exceptions in lambda automatically become `OperationResult { Success = false }`
- Double-wrapping (try-catch returning error result) loses stack context and originates from wrong layer
- Finally blocks are the correct place for resource cleanup, NOT catch blocks for error suppression

**See:** CRITICAL-RULES.md Rule 1 for Success flag requirements

---

## MCP Server Resource-Based Tools

**In-Process Architecture**: MCP Server hosts ExcelMcpService fully in-process with direct method calls (no pipe).
ServiceBridge holds the service reference and calls ProcessAsync() directly.

**19 Focused Tools:**
1. `file` - Session lifecycle (open, close, create, list)
2. `worksheet` - Worksheet operations
3. `worksheet_style` - Tab colors and visibility
4. `range` - Range values and formulas
5. `range_edit` - Insert/delete/find/replace
6. `table` - Excel Tables (ListObjects)
7. `table_column` - Table columns/filters/sorts
8. `powerquery` - Power Query M code
9. `pivottable` - PivotTable lifecycle
10. `pivottable_field` - PivotTable fields
11. `pivottable_calc` - Calculated fields/items
12. `chart` - Chart lifecycle
13. `chart_config` - Chart configuration
14. `connection` - Data connections
15. `slicer` - Slicers
16. `vba` - VBA macros
17. `datamodel` - Power Pivot / DAX
18. `datamodel_relationship` - Data Model relationships
19. `namedrange` - Named ranges
20. `excel_calculation` - Calculation mode

### Action-Based Routing with ForwardToService
```csharp
[McpServerTool]
public static string ExcelPowerQuery(string action, string sessionId, ...)
{
    return action.ToLowerInvariant() switch
    {
        "list" => ForwardList(sessionId),
        "view" => ForwardView(sessionId, queryName),
        _ => throw new McpException($"Unknown action: {action}")
    };
}

private static string ForwardList(string sessionId)
{
    return ExcelToolsBase.ForwardToService("powerquery.list", sessionId);
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

**Minimize workbook opens** - Use single session for multiple operations
**Bulk operations** - Use `range.Value2` for 2D arrays, not cell-by-cell access

---

## Key Principles

1. **WithExcel() for everything** - See excel-com-interop.instructions.md
2. **Release intermediate objects** - Prevents Excel hanging
3. **Batch/Session for MCP** - Multiple operations in single session
4. **Resource-based tools** - 22 tools, not 33+ operations
5. **DRY utilities** - Share common patterns
6. **Security defaults** - Never expose credentials
7. **Bulk operations** - Minimize COM round-trips
