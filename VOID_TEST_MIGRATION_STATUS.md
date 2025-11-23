# Void Test Migration Progress

## Status: 22/84 Errors Fixed (26%) ✅

### Completed ✅ (22 errors)
- **SheetCommandsTests.Lifecycle.cs** (2 errors): Create, Copy
- **SheetCommandsTests.Move.cs** (11 errors): Move, CopyToWorkbook x4, MoveToWorkbook x3, Copy
- **SheetCommandsTests.TabColor.cs** (2 errors): SetTabColor, ClearTabColor  
- **SheetCommandsTests.Visibility.cs** (6 errors): SetVisibility, Show, Hide, VeryHide
- **Helper files** (4 errors fixed earlier): PowerQuery, Table, DataModel, PivotTable fixtures

### Build Status: 62 errors remaining

## Remaining Errors by Category

### HIGH PRIORITY: PowerQueryCommandsTests (16 errors)
**Type**: CS0815 - "Cannot assign void"  
**Pattern**: `var result = _powerQueryCommands.Update/Delete/LoadTo/RefreshAll(...)`  
**Fix**: Remove `var result =` prefix, remove `Assert.True(result.Success...)` line

Lines needing fixes:
- 208, 212, 363, 424, 434, 518, 528, 592, 602, 665, 688, 711, 781, 803, 807, 810, 813

Example fix:
```csharp
// BEFORE
var result = _powerQueryCommands.Update(batch, queryName, filePath);
Assert.True(result.Success, $"Update failed: {result.ErrorMessage}");

// AFTER
_powerQueryCommands.Update(batch, queryName, filePath);  // Update throws on error
```

### HIGH PRIORITY: DataModelCommandsTests (12 errors)
**Type**: CS4008 - "Cannot await void"  
**Pattern**: `await _dataModelCommands.CreateMeasure/DeleteMeasure/etc(...)`  
**Fix**: Remove `await` keyword OR convert async test method to non-async

Lines: 158, 181, 185, 205, 209, 267, 271, 293, 305

Example fix:
```csharp
// BEFORE
await _dataModelCommands.CreateMeasure(batch, tableName, measureName, formula);

// AFTER
_dataModelCommands.CreateMeasure(batch, tableName, measureName, formula);  // CreateMeasure throws on error
```

**Note**: Tests are currently async (`public async Task`). Two options:
1. Remove `await` keywords from void method calls (simpler, 1-line fixes per error)
2. Convert test methods to non-async (complex, multiple changes per test)

**Recommended**: Option 1 - just remove the `await` keywords

### MEDIUM PRIORITY: DataModelTestsFixture (1 error)
**Type**: CS0815 - "Cannot assign void"  
**Line**: 91

### MEDIUM PRIORITY: TableCommandsTests (1 error)  
**Type**: CS0815 - "Cannot assign void"  
**Line**: 400

### LOW PRIORITY: PivotTableCommandsTests (2 errors)
**Type**: CS0815 - "Cannot assign void"  
**Lines**: 47, 182, 228

## Next Steps

1. **Fix PowerQueryCommandsTests** (16 fixes): Most straightforward - just remove `var result =`
2. **Fix DataModelCommandsTests** (12 fixes): Remove `await` keywords from 12 lines
3. **Fix remaining helper/command files** (4 fixes): Follow same pattern as completed tests
4. **Verify clean build**: `dotnet build` should show 0 errors
5. **Run smoke test**: `dotnet test --filter "SmokeTest_AllTools_LlmWorkflow"`

## Systematic Fix Pattern

All remaining fixes follow this simple pattern:

```csharp
// PATTERN 1: Remove var assignment (most common)
// BEFORE
var result = _commands.VoidMethod(batch, args);
Assert.True(result.Success, $"Method failed: {result.ErrorMessage}");

// AFTER  
_commands.VoidMethod(batch, args);  // VoidMethod throws on error

// PATTERN 2: Remove await from void methods
// BEFORE
await _commands.VoidMethod(batch, args);

// AFTER
_commands.VoidMethod(batch, args);  // VoidMethod throws on error
```

## Architecture Context

- **Void Execute Infrastructure**: Added to IExcelBatch, enables clean void operations for mutation commands
- **Exception Model**: Commands throw InvalidOperationException on error (from remote branch)
- **Test Updates Required**: Only test layer affected - Core/CLI/MCP layers already support exception-first pattern
- **No API Changes**: void Execute just calls ExecuteTask<int>(...) with dummy 0 return
