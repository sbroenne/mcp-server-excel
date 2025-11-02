---
applyTo: "tests/**/*.cs"
---

# Testing Strategy - Quick Reference

> **Two tiers: Integration (Excel) → OnDemand (session cleanup)**
> 
> **⚠️ No Unit Tests**: ExcelMcp has no traditional unit tests. Integration tests ARE our unit tests because Excel COM cannot be meaningfully mocked. See `docs/ADR-001-NO-UNIT-TESTS.md` for full rationale.

## Test Class Templates

### Integration Test Template

```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "FeatureName")]  // PowerQuery, DataModel, Tables, PivotTables, Ranges, Connections, Parameters, Worksheets
[Trait("RequiresExcel", "true")]
public partial class FeatureCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly IFeatureCommands _commands;
    private readonly string _tempDir;

    public FeatureCommandsTests(TempDirectoryFixture fixture)
    {
        _commands = new FeatureCommands();
        _tempDir = fixture.TempDir;
    }

    [Fact]
    public async Task Operation_Scenario_ExpectedResult()
    {
        // Arrange - Each test gets unique file
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(FeatureCommandsTests), 
            nameof(Operation_Scenario_ExpectedResult), 
            _tempDir,
            ".xlsx");  // Use ".xlsm" for VBA tests

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _commands.OperationAsync(batch, args);

        // Assert - Verify actual Excel state, not just success flag
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        
        // Verify object exists/updated in Excel (round-trip validation)
        var verifyResult = await _commands.ListAsync(batch);
        Assert.Contains(verifyResult.Items, i => i.Name == "Expected");
        
        // No SaveAsync unless testing persistence (see examples below)
    }
}
```

## Essential Rules

### File Isolation
- ✅ Each test creates unique file via `CoreTestHelper.CreateUniqueTestFileAsync()`
- ❌ **NEVER** share test files between tests
- ✅ Use `.xlsm` for VBA tests, `.xlsx` otherwise

### Assertions  
- ✅ Binary assertions: `Assert.True(result.Success, $"Reason: {result.ErrorMessage}")`
- ❌ **NEVER** "accept both" patterns
- ✅ **ALWAYS verify actual Excel state** after create/update operations

### SaveAsync
- ❌ **FORBIDDEN** unless explicitly testing persistence
- ✅ **ONLY** for round-trip tests: Create → Save → Re-open → Verify
- ❌ **NEVER** call in middle of test (breaks subsequent operations)
- See CRITICAL-RULES.md Rule 14 for details

**Examples:**

```csharp
// ❌ WRONG: SaveAsync in middle breaks next operation
await _commands.CreateAsync(batch, "Sheet1");
await batch.SaveAsync();  // ❌ Breaks subsequent operations!
await _commands.RenameAsync(batch, "Sheet1", "New");  // FAILS!

// ✅ CORRECT: SaveAsync only at end
await _commands.CreateAsync(batch, "Sheet1");
await _commands.RenameAsync(batch, "Sheet1", "New");
await batch.SaveAsync();  // ✅ After all operations

// ✅ CORRECT: Persistence test with re-open
await using var batch1 = await ExcelSession.BeginBatchAsync(testFile);
await _commands.CreateAsync(batch1, "Sheet1");
await batch1.SaveAsync();  // Save for persistence

await using var batch2 = await ExcelSession.BeginBatchAsync(testFile);
var list = await _commands.ListAsync(batch2);
Assert.Contains(list.Items, i => i.Name == "Sheet1");  // ✅ Verify persisted
```

### Batch Pattern
- All Core commands accept `IExcelBatch batch` as first parameter
- Use `await using var batch` for automatic disposal
- **NEVER** call `SaveAsync()` in middle of test

### Required Traits
- `[Trait("Category", "Integration")]` - All tests are integration tests
- `[Trait("Speed", "Medium|Slow")]`
- `[Trait("Layer", "Core|CLI|McpServer|ComInterop")]`
- `[Trait("Feature", "<feature-name>")]` - See valid values below
- `[Trait("RequiresExcel", "true")]` - All integration tests require Excel
- `[Trait("RunType", "OnDemand")]` - For session/lifecycle tests only

### Valid Feature Values
- **PowerQuery** - Power Query M code operations
- **DataModel** - Data Model / DAX operations
- **Tables** - Excel Table (ListObject) operations
- **PivotTables** - PivotTable operations
- **Ranges** - Range data operations
- **Connections** - Connection management
- **Parameters** - Named range parameters
- **Worksheets** - Worksheet lifecycle
- **VBA** - VBA script operations
- **VBATrust** - VBA trust detection/configuration

## Test Execution

```bash
# Development (fast feedback - excludes VBA tests)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Pre-commit (comprehensive - excludes VBA tests)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Session/batch changes (MANDATORY - see CRITICAL-RULES.md Rule 3)
dotnet test --filter "RunType=OnDemand"

# VBA tests (run manually when needed - requires VBA trust enabled)
dotnet test --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"
```

## Round-Trip Validation Pattern

**Always verify actual Excel state after operations:**

```csharp
// ✅ CREATE → Verify exists
var createResult = await _commands.CreateAsync(batch, "TestTable");
Assert.True(createResult.Success);

var listResult = await _commands.ListAsync(batch);
Assert.Contains(listResult.Items, i => i.Name == "TestTable");  // ✅ Proves it exists!

// ✅ UPDATE → Verify changes applied
var updateResult = await _commands.RenameAsync(batch, "TestTable", "NewName");
Assert.True(updateResult.Success);

var viewResult = await _commands.GetAsync(batch, "NewName");
Assert.Equal("NewName", viewResult.Name);  // ✅ Proves rename worked!

// ✅ DELETE → Verify removed
var deleteResult = await _commands.DeleteAsync(batch, "NewName");
Assert.True(deleteResult.Success);

var finalList = await _commands.ListAsync(batch);
Assert.DoesNotContain(finalList.Items, i => i.Name == "NewName");  // ✅ Proves deletion!
```

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| Shared test file | Each test creates unique file |
| Only test success flag | Verify actual Excel state |
| SaveAsync before assertions | Remove SaveAsync entirely |
| SaveAsync in middle of test | Only at end or in persistence test |
| Manual IDisposable | Use `IClassFixture<TempDirectoryFixture>` |
| .xlsx for VBA tests | Use `.xlsm` |
| "Accept both" assertions | Binary assertions only |
| Missing Feature trait | Add from valid feature list above |

## When Tests Fail

1. Run individually: `--filter "FullyQualifiedName=Namespace.Class.Method"`
2. Check file isolation (unique files?)
3. Check assertions (binary, not conditional?)
4. Check SaveAsync (removed unless persistence test?)
5. Verify Excel state (not just success flag?)

**Full checklist**: See CRITICAL-RULES.md Rule 12
