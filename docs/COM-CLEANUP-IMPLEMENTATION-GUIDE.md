# COM Object Cleanup Implementation Guide

## Overview

This guide documents the systematic approach for implementing COM object cleanup across all command methods in the ExcelMcp codebase. This work is part of addressing user-reported issues with GC warnings and Excel processes not terminating properly.

## Pattern to Apply

### Basic Pattern

For any method that uses COM objects from Excel:

```csharp
public OperationResult SomeMethod(string filePath, string param)
{
    var result = new OperationResult { /* init */ };
    
    WithExcel(filePath, save, (excel, workbook) =>
    {
        dynamic? comObject1 = null;
        dynamic? comObject2 = null;
        try
        {
            comObject1 = workbook.SomeCollection;
            comObject2 = comObject1.Item(1);
            
            // Use COM objects...
            
            result.Success = true;
            return 0;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
            return 1;
        }
        finally
        {
            // Release in reverse order of creation
            ReleaseComObject(ref comObject2);
            ReleaseComObject(ref comObject1);
        }
    });
    
    return result;
}
```

### Collection Iteration Pattern

When iterating through COM collections:

```csharp
dynamic? collection = null;
try
{
    collection = workbook.Items;
    
    for (int i = 1; i <= collection.Count; i++)
    {
        dynamic? item = null;
        try
        {
            item = collection.Item(i);
            // Use item...
        }
        finally
        {
            ReleaseComObject(ref item);
        }
    }
}
finally
{
    ReleaseComObject(ref collection);
}
```

### Keeping References Pattern

When you need to keep a COM object to return or use later:

```csharp
dynamic? target = null;

for (int i = 1; i <= collection.Count; i++)
{
    dynamic? item = null;
    try
    {
        item = collection.Item(i);
        if (/* match condition */)
        {
            target = item;
            item = null; // Don't release - we're keeping it
            break;
        }
    }
    finally
    {
        ReleaseComObject(ref item); // Only releases non-null items
    }
}

// Use target...
ReleaseComObject(ref target); // Release when done
```

## Methods Completed

### PowerQueryCommands.cs (11/18 = 61%)

**Completed**:
1. DetectPrivacyLevelsAndRecommend() - Phase 1
2. ApplyPrivacyLevel() - Phase 1
3. DetermineLoadedSheet() - Phase 1
4. List() - Phase 1
5. View() - Phase 2
6. Update() - Phase 3
7. Export() - Phase 3
8. Import() - Phase 3
9. Refresh() - Phase 5
10. Errors() - Phase 5
11. LoadTo() - Phases 5-6
12. Delete() - Phase 6
13. SetConnectionOnly() - Phase 6

**Remaining** (8 methods):
- Sources() - line 1517
- Test() - line 1585
- Peek() - line 1646
- Eval() - line 1734
- SetLoadToTable() - line 1841
- SetLoadToDataModel() - line 1925
- SetLoadToBoth() - line 2054
- GetLoadConfig() - line 2194

### ExcelHelper.cs (3/3 = 100%)

**Completed**:
1. RemoveQueryConnections() - Phase 4
2. GetQueryNames() - Phase 4
3. FindQuery() - Phase 4

## Methods Remaining

### ScriptCommands.cs (0/6 = 0%)

All 6 methods need COM cleanup:
1. List() - VBProject, VBComponents iteration
2. Export() - VBProject, VBComponents access
3. Import() - VBProject, VBComponents, CodeModule
4. Update() - VBProject, VBComponents, CodeModule
5. Run() - VBProject, VBComponents, Application.Run
6. Delete() - VBProject, VBComponents

**Pattern**: Always release VBProject, VBComponents collection, individual components, and CodeModule objects.

### CellCommands.cs (0/4 = 0%)

All 4 methods need COM cleanup:
1. GetValue() - Sheet, Cell/Range objects
2. SetValue() - Sheet, Cell/Range objects
3. GetFormula() - Sheet, Cell/Range objects
4. SetFormula() - Sheet, Cell/Range objects

**Pattern**: Always release Sheet and Range objects.

### ParameterCommands.cs (0/5 = 0%)

All 5 methods need COM cleanup:
1. List() - Names collection, Name objects
2. Get() - Names collection, Name object, RefersToRange
3. Set() - Names collection, Name object, RefersToRange
4. Create() - Names collection
5. Delete() - Names collection, Name object

**Pattern**: Always release Names collection, individual Name objects, and RefersToRange objects.

### SheetCommands.cs (Core) (0/9 = 0%)

All 9 methods need COM cleanup:
1. List() - Worksheets collection, individual sheets
2. Read() - Sheet, Range objects
3. Write() - Sheet, Range, Cells objects
4. Create() - Worksheets collection
5. Rename() - Sheet object
6. Copy() - Worksheets collection, source and target sheets
7. Delete() - Worksheets collection, Sheet object
8. Clear() - Sheet, Range objects
9. Append() - Sheet, Range, used range objects

**Pattern**: Always release Worksheets collection, individual Sheet objects, and Range objects.

### SetupCommands.cs (0/3 = 0%)

All 3 methods need COM cleanup:
1. CheckVbaTrust() - VBProject access
2. EnableVbaTrust() - Registry operations (no COM)
3. DisableVbaTrust() - Registry operations (no COM)

**Pattern**: Only first method needs COM cleanup for VBProject.

### FileCommands.cs (0/1 = 0%)

1. CreateEmpty() - No COM objects used during creation (uses WithNewExcel helper which handles cleanup)

**Pattern**: No changes needed - already handled by WithNewExcel.

## Implementation Steps for Remaining Methods

For each remaining method:

1. **Identify all COM object creations**:
   - Collections: `workbook.Queries`, `workbook.Worksheets`, `workbook.Names`, etc.
   - Items: `collection.Item(i)`, `sheet.Range["A1"]`, etc.
   - Properties: `nameObj.RefersToRange`, `query.Formula`, etc.

2. **Declare as nullable dynamic at appropriate scope**:
   ```csharp
   dynamic? comObject = null;
   ```

3. **Wrap usage in try-finally**:
   ```csharp
   try
   {
       comObject = /* assignment */;
       // use it
   }
   finally
   {
       ReleaseComObject(ref comObject);
   }
   ```

4. **Handle loops carefully**:
   - Loop variable COM objects inside loop try-finally
   - Collection COM object outside loop try-finally

5. **Test after each batch**:
   ```bash
   dotnet build -c Release
   dotnet test --no-build -c Release
   ```

## Common Mistakes to Avoid

1. **Not declaring as nullable**: Must use `dynamic?` not `dynamic`
2. **Wrong try-finally order**: catch before finally, not after
3. **Forgetting loop items**: Each `collection.Item(i)` must be released
4. **Not releasing intermediate objects**: Even temporary objects need cleanup
5. **Releasing objects still in use**: Use `item = null` to prevent release when keeping reference

## Benefits Achieved

- ✅ Eliminates GC pressure warnings
- ✅ Prevents Excel processes from lingering
- ✅ Reduces memory footprint
- ✅ Improves reliability of COM interop
- ✅ Follows Microsoft recommended patterns

## Testing Strategy

After implementing COM cleanup:

1. **Build**: Must succeed with 0 warnings
2. **Unit Tests**: All 74 tests must pass
3. **Manual verification**: 
   - Run command
   - Wait 5 seconds
   - Check Task Manager for excel.exe
   - Should be no Excel processes running

## References

- Microsoft COM Interop Best Practices
- `ExcelHelper.ReleaseComObject<T>()` method documentation
- `.github/copilot-instructions.md` - COM cleanup section
