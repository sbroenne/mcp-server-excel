# Core Test Refactoring - Completion Guide

## Status: 4 of 7 Complete ✅

### Completed Test Classes
1. ✅ **ParameterCommandsTests** → `Parameter/` (3 files)
2. ✅ **SheetCommandsTests** → `Sheet/` (2 files)
3. ✅ **SetupCommandsTests** → `Setup/` (2 files)
4. ✅ **ScriptCommandsTests** → `Script/` (2 files)

### Remaining Test Classes (Large Files)
5. ⏳ **TableCommandsTests** (509 lines) → `Table/` (4 files needed)
6. ⏳ **PivotTableCommandsTests** (660+ lines) → `PivotTable/` (4 files needed)
7. ⏳ **ConnectionCommandsTests** (575+ lines) → `Connection/` (5 files needed)

## Refactoring Pattern (Proven)

### Step 1: Create Directory Structure
```bash
mkdir -p tests/ExcelMcp.Core.Tests/Integration/Commands/{Table,PivotTable,Connection}
```

### Step 2: Create Base Partial Class

**Template:** `{CommandName}Tests.cs`
```csharp
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.{CommandName};

[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "{FeatureName}")]
public partial class {CommandName}Tests : IDisposable
{
    private readonly I{CommandName} _commands;
    private readonly string _tempDir;
    private bool _disposed;

    public {CommandName}Tests()
    {
        _commands = new {CommandName}();
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_{CommandName}_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (_disposed) return;
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch { }
        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
```

### Step 3: Create Partial Files by Feature

Each test method becomes:
```csharp
[Fact]
public async Task {TestName}()
{
    // Arrange - Create unique file
    var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
        nameof({CommandName}Tests), nameof({TestName}), _tempDir);

    // Act - Use single batch for all related operations
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);
    
    // Setup test data if needed
    await SetupTestData(batch);
    
    // Perform operation
    var result = await _commands.{Operation}Async(batch, args);
    
    // Assert
    Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
    
    // Verify changes occurred (CRUD validation)
    var verifyResult = await _commands.List/GetAsync(batch, ...);
    Assert.Contains/DoesNotContain(...);
    
    await batch.SaveAsync();
}
```

### Step 4: Key Refactoring Rules

1. **File Isolation:** Each test creates unique file via `CoreTestHelper.CreateUniqueTestFileAsync()`
2. **Single Batch:** Combine setup → operation → verify in ONE batch
3. **CRUD Validation:** Create/Update → verify with List/Get, Delete → verify with List
4. **Remove Duplicates:** Merge tests like "Create" and "Create_ThenList" into single test
5. **Remove Shared State:** No `_testExcelFile` field

### Step 5: Batching Optimization Examples

❌ **Before (Multiple Batches):**
```csharp
await using (var batch = ...) { await CreateAsync(...); await batch.SaveAsync(); }
await using (var batch = ...) { await ListAsync(...); }  // 2nd open!
```

✅ **After (Single Batch):**
```csharp
await using var batch = await ExcelSession.BeginBatchAsync(testFile);
await _commands.CreateAsync(batch, ...);
var listResult = await _commands.ListAsync(batch);  // Same batch!
await batch.SaveAsync();
```

## Remaining Work Details

### TableCommandsTests (509 lines → 4 files)

**Regions identified:**
1. `TableCommandsTests.Lifecycle.cs` - List, Create, Info (Phase 1)
2. `TableCommandsTests.StructuredReferences.cs` - GetStructuredReference tests (Phase 2)
3. `TableCommandsTests.Sorting.cs` - Sort tests (Phase 2)
4. `TableCommandsTests.DataModel.cs` - AddToDataModel tests

**Key fixes needed:**
- Remove shared `_testExcelFile`
- Each test creates unique file with table data
- Helper method to create table: `CreateTestFileWithTableAsync(fileName)`

### PivotTableCommandsTests (660+ lines → 4 files)

**Suggested split:**
1. `PivotTableCommandsTests.Creation.cs` - CreateFromRange, CreateFromTable
2. `PivotTableCommandsTests.Fields.cs` - Field operations
3. `PivotTableCommandsTests.Layout.cs` - Layout operations
4. `PivotTableCommandsTests.Filters.cs` - Filter operations

**Key fixes needed:**
- Uses helper `CreateTestFileWithDataAsync(fileName)` - keep this pattern
- Already creates unique file per test - good!
- Just needs directory organization

### ConnectionCommandsTests (575+ lines → 5 files)

**Suggested split:**
1. `ConnectionCommandsTests.List.cs` - List operations
2. `ConnectionCommandsTests.View.cs` - View/GetProperties operations
3. `ConnectionCommandsTests.Import.cs` - Import/Create operations
4. `ConnectionCommandsTests.Update.cs` - Update/Refresh operations
5. `ConnectionCommandsTests.Properties.cs` - Property get/set operations

**Key fixes needed:**
- Remove shared `_testFile`, `_testCsvFile`
- Each test creates unique CSV file if needed
- Uses `ConnectionTestHelper` - keep using it

## Testing After Refactoring

```bash
# Build
dotnet build tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj

# Run specific test class
dotnet test --filter "FullyQualifiedName~TableCommandsTests"

# Run all Core integration tests
dotnet test --filter "(Category=Integration)&RunType!=OnDemand" tests/ExcelMcp.Core.Tests
```

## Success Criteria

- [ ] All 7 test classes refactored into partial files in directories
- [ ] Every test method uses `CoreTestHelper.CreateUniqueTestFileAsync()`
- [ ] No shared `_testExcelFile` fields
- [ ] Single batch for related operations
- [ ] CRUD tests verify changes occurred
- [ ] No duplicate tests
- [ ] All tests pass
- [ ] Build succeeds with 0 warnings

## References

**Completed Examples:**
- `tests/ExcelMcp.Core.Tests/Integration/Commands/Parameter/` (3 files)
- `tests/ExcelMcp.Core.Tests/Integration/Commands/Sheet/` (2 files)
- `tests/ExcelMcp.Core.Tests/Integration/Commands/Setup/` (2 files)
- `tests/ExcelMcp.Core.Tests/Integration/Commands/Script/` (2 files)

**Good Existing Examples:**
- `tests/ExcelMcp.Core.Tests/Integration/Range/` (7 partial files)
- `tests/ExcelMcp.Core.Tests/Integration/Commands/DataModel/` (5 partial files)
