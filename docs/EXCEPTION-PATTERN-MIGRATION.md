# Exception Pattern Migration - Test Fixes Required

> **Status**: 82 test compilation errors (fixed 2 of 84, 97.6% remaining)  
> **Root Cause**: Commands converted to throw exceptions (void methods) but tests still use `var result = ` pattern  
> **Solution**: Systematically remove variable assignments for void method calls

## Progress Summary

**Completed**:
- ✅ void Execute infrastructure added to IExcelBatch and ExcelBatch (non-breaking)
- ✅ Documented all 84 errors with clear categorization
- ✅ Fixed 2/84 errors (PowerQueryCommandsTests.cs: 4 fixes → 2 remaining)
- ✅ Fixed helper fixtures (PowerQueryTestsFixture, TableTestsFixture, DataModelTestsFixture, PivotTableRealisticFixture)
- ✅ Fixed FileCommandsTests.CreateEmptyThenList
- ✅ Fixed SheetCommandsTests.Lifecycle (all 4 errors)
- ✅ Fixed TableCommandsTests.cs (14 fixes - all void Create/Delete/Rename/etc patterns)

**Remaining**: 82 errors across:
- SheetCommandsTests.Move.cs (10 errors - Move, CopyToWorkbook, MoveToWorkbook calls)
- SheetCommandsTests.TabColor.cs (2 errors)
- SheetCommandsTests.Visibility.cs (6 errors)
- PowerQueryCommandsTests.cs (16 remaining errors)
- DataModelCommandsTests.cs (8 remaining errors - await void patterns)
- PivotTableCommandsTests.Creation.cs (1 error)
- PivotTableCommandsTests.OlapFields.cs (2 errors)

## Error Summary

Total errors: **82** (down from 84)

### Error Types

| Error Type | Count | Pattern | Fix |
|-----------|-------|---------|-----|
| CS0815 | 75 | `var result = _command.VoidMethod()` | Remove `var result = ` |
| CS4008 | 9 | `await _command.VoidMethod()` | Remove `await` |

### Affected Test Files

| File | Line Count | Errors | Status |
|------|-----------|--------|--------|
| TableCommandsTests.cs | 500+ | 16 | 🔴 Pending |
| SheetCommandsTests.Lifecycle.cs | 150+ | 4 | 🔴 Pending |
| SheetCommandsTests.Move.cs | 370+ | 10 | 🔴 Pending |
| SheetCommandsTests.TabColor.cs | 140+ | 2 | 🔴 Pending |
| SheetCommandsTests.Visibility.cs | 180+ | 6 | 🔴 Pending |
| PowerQueryCommandsTests.cs | 850+ | 20 | 🔴 Pending |
| DataModelCommandsTests.cs | 320+ | 8 | 🔴 Pending |
| PivotTableCommandsTests.Creation.cs | 80+ | 1 | 🔴 Pending |
| PivotTableCommandsTests.OlapFields.cs | 250+ | 2 | 🔴 Pending |
| PowerQueryTestsFixture.cs (helper) | 100+ | 3 | 🔴 Pending |
| DataModelTestsFixture.cs (helper) | 150+ | 8 | 🔴 Pending |
| PivotTableRealisticFixture.cs (helper) | 150+ | 2 | 🔴 Pending |
| TableTestsFixture.cs (helper) | 120+ | 1 | 🔴 Pending |
| FileCommandsTests.CreateEmptyThenList.cs | 100+ | 1 | 🔴 Pending |

## Fix Pattern

### Pattern 1: Simple Void Call (CS0815)

**Before:**
```csharp
var result = await _commands.CreateAsync(batch, "Name");
Assert.NotNull(result); // ❌ result is void
```

**After:**
```csharp
await _commands.CreateAsync(batch, "Name");
// Exceptions throw if operation fails
```

### Pattern 2: Void Call Without Await (CS0815)

**Before:**
```csharp
var result = _testHelper.SetupTable(batch, data);
Assert.NotNull(result);
```

**After:**
```csharp
_testHelper.SetupTable(batch, data);
// Exceptions throw if operation fails
```

### Pattern 3: Void with Await (CS4008)

**Before:**
```csharp
var result = await _commands.DeleteAsync(batch, name);
Assert.True(result.Success); // ❌ Can't await void
```

**After:**
```csharp
await _commands.DeleteAsync(batch, name);
// Exceptions throw if operation fails
```

## Verification Strategy

**Per-file approach:**
1. Fix `var result = ` assignments to void methods
2. Remove `Assert.NotNull(result)`, `Assert.True(result.Success)` checks (void methods throw on error)
3. Keep assertions that verify actual Excel state (round-trip validation)
4. Run `dotnet build` to confirm file's errors resolved
5. Move to next file

**Build validation:**
```bash
dotnet build  # Must complete with 0 errors
```

## Commands for Bulk Search

Find all `var result = ` assignments:
```bash
git grep -n "var result = " tests/ExcelMcp.Core.Tests/
```

Find all void method calls with assignment:
```bash
git grep -n "var .* = .*\..*Async(batch"
```

## Implementation Priority

**Phase 1 (High Error Count):**
1. PowerQueryCommandsTests.cs - 20 errors
2. TableCommandsTests.cs - 16 errors
3. SheetCommandsTests.Move.cs - 10 errors
4. DataModelCommandsTests.cs - 8 errors + DataModelTestsFixture (8 errors)

**Phase 2 (Helpers):**
5. PowerQueryTestsFixture.cs - 3 errors
6. PivotTableRealisticFixture.cs - 2 errors
7. TableTestsFixture.cs - 1 error

**Phase 3 (Remaining):**
8. SheetCommandsTests.Lifecycle/TabColor/Visibility - 12 errors combined
9. PivotTableCommandsTests - 3 errors
10. FileCommandsTests.CreateEmptyThenList.cs - 1 error

## Testing After Fixes

Once all 84 errors are fixed:

```bash
# Build to verify compilation
dotnet build

# Run smoke test
dotnet test --filter "FullyQualifiedName~McpServerSmokeTests.SmokeTest_AllTools_LlmWorkflow" --verbosity quiet

# Run feature-specific tests
dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"
dotnet test --filter "Feature=Tables&RunType!=OnDemand"
```

## Notes

- **No API changes needed** - void Execute infrastructure is complete
- **Just test updates** - All test files need removal of `var result = ` patterns
- **Error-first semantics** - Commands throw on error, tests no longer check `Success` flag
- **Exception handling** - Try-catch in tests if specific error handling is needed (rare)

## Rollback Procedure

If needed:
```bash
git reset --hard HEAD~1  # Undo void Execute commit
git checkout HEAD -- src/ExcelMcp.ComInterop/Session/IExcelBatch.cs
git checkout HEAD -- src/ExcelMcp.ComInterop/Session/ExcelBatch.cs
```
