---
applyTo: "tests/**/*.cs"
---

# Testing Strategy - Quick Reference

> **Three tiers: Unit (fast) → Integration (Excel) → OnDemand (session cleanup)**

## Test Class Template

```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "FeatureName")]
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
            ".xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _commands.OperationAsync(batch, args);

        // Assert - Verify actual Excel state, not just success flag
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        
        // Verify object exists/updated in Excel
        var verifyResult = await _commands.ListAsync(batch);
        Assert.Contains(verifyResult.Items, i => i.Name == "Expected");
        
        // No SaveAsync unless testing persistence (round-trip)
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
- See CRITICAL-RULES.md Rule 14 for details

### Batch Pattern
- All Core commands accept `IExcelBatch batch` as first parameter
- Use `await using var batch` for automatic disposal
- **NEVER** call `SaveAsync()` in middle of test

### Required Traits
- `[Trait("Category", "Integration|Unit")]`
- `[Trait("Speed", "Medium|Fast")]`
- `[Trait("Layer", "Core|CLI|McpServer")]`
- `[Trait("RequiresExcel", "true")]` for integration tests

## Test Execution

```bash
# Development (fast feedback)
dotnet test --filter "Category=Unit&RunType!=OnDemand"

# Pre-commit (comprehensive)
dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand"

# Session/batch changes (MANDATORY - see CRITICAL-RULES.md Rule 3)
dotnet test --filter "RunType=OnDemand"
```

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| Shared test file | Each test creates unique file |
| Only test success flag | Verify actual Excel state |
| SaveAsync before assertions | Remove SaveAsync entirely |
| Manual IDisposable | Use `IClassFixture<TempDirectoryFixture>` |
| .xlsx for VBA tests | Use `.xlsm` |
| "Accept both" assertions | Binary assertions only |

## When Tests Fail

1. Run individually: `--filter "FullyQualifiedName=Namespace.Class.Method"`
2. Check file isolation (unique files?)
3. Check assertions (binary, not conditional?)
4. Check SaveAsync (removed unless persistence test?)
5. Verify Excel state (not just success flag?)

**Full checklist**: See CRITICAL-RULES.md Rule 12
