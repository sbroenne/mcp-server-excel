---
applyTo: "tests/**/*.cs"
---

# Testing Strategy - Quick Reference

> **Two tiers: Integration (Excel) → OnDemand (session cleanup)**
> 
> **⚠️ No Unit Tests**: ExcelMcp has no traditional unit tests. Integration tests ARE our unit tests because Excel COM cannot be meaningfully mocked. See `docs/ADR-001-NO-UNIT-TESTS.md` for full rationale.

## Test Naming Standard

**Pattern**: `MethodName_StateUnderTest_ExpectedBehavior`

- **MethodName**: Command being tested (no "Async" suffix)
- **StateUnderTest**: Specific scenario/condition (not generic like "Valid")
- **ExpectedBehavior**: Clear outcome (Returns*, Creates*, Removes*, etc.)

**Examples**:
```csharp
✅ List_EmptyWorkbook_ReturnsEmptyList
✅ Create_UniqueName_ReturnsSuccess
✅ Delete_NonActiveSheet_ReturnsSuccess
✅ ImportThenDelete_UniqueQuery_RemovedFromList

❌ List_WithValidFile_ReturnsSuccessResult  // Too generic
❌ CreateAsync_ValidName_Success            // Has Async suffix
❌ Delete                                   // Missing state and behavior
```

**Full Standard**: See `docs/TEST-NAMING-STANDARD.md` for complete guide with pattern catalog and examples.

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
        var testFile = CoreTestHelper.CreateUniqueTestFile(
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
        
        // No Save unless testing persistence (see examples below)
    }
}
```

## Essential Rules

### File Isolation
- ✅ Each test creates unique file via `CoreTestHelper.CreateUniqueTestFile()`
- ❌ **NEVER** share test files between tests
- ✅ Use `.xlsm` for VBA tests, `.xlsx` otherwise

### Assertions  
- ✅ Binary assertions: `Assert.True(result.Success, $"Reason: {result.ErrorMessage}")`
- ❌ **NEVER** "accept both" patterns
- ✅ **ALWAYS verify actual Excel state** after create/update operations

### Diagnostic Output
- ✅ Use `ILogger` with `ITestOutputHelper` for diagnostic messages
- ✅ Pattern: Test constructor receives `ITestOutputHelper`, creates logger via `MartinCostello.Logging.XUnit`
- ✅ Pass logger to `ExcelBatch` constructor (requires `InternalsVisibleTo` for accessing internal ExcelBatch)
- ✅ Logger messages appear in test output automatically (success or failure)
- ❌ **NEVER** use `Console.WriteLine()` - output is suppressed by test runner
- ❌ **NEVER** use `Debug.WriteLine()` - only visible with debugger attached, not in test output
- ❌ **NEVER** write to files for diagnostics - use proper logging infrastructure

### Save
- ❌ **FORBIDDEN** unless explicitly testing persistence
- ✅ **ONLY** for round-trip tests: Create → Save → Re-open → Verify
- ❌ **NEVER** call in middle of test (breaks subsequent operations)
- See CRITICAL-RULES.md Rule 14 for details

**Examples:**

```csharp
// ❌ WRONG: Save in middle breaks next operation
await _commands.CreateAsync(batch, "Sheet1");
await batch.Save();  // ❌ Breaks subsequent operations!
await _commands.RenameAsync(batch, "Sheet1", "New");  // FAILS!

// ✅ CORRECT: Save only at end
await _commands.CreateAsync(batch, "Sheet1");
await _commands.RenameAsync(batch, "Sheet1", "New");
await batch.Save();  // ✅ After all operations

// ✅ CORRECT: Persistence test with re-open
await using var batch1 = await ExcelSession.BeginBatchAsync(testFile);
await _commands.CreateAsync(batch1, "Sheet1");
await batch1.Save();  // Save for persistence

await using var batch2 = await ExcelSession.BeginBatchAsync(testFile);
var list = await _commands.ListAsync(batch2);
Assert.Contains(list.Items, i => i.Name == "Sheet1");  // ✅ Verify persisted
```

### Batch Pattern
- All Core commands accept `IExcelBatch batch` as first parameter
- Use `await using var batch` for automatic disposal
- **NEVER** call `Save()` in middle of test

### Required Traits
- `[Trait("Category", "Integration")]` - All tests are integration tests
- `[Trait("Speed", "Medium|Slow")]`
- `[Trait("Layer", "Core|CLI|McpServer|ComInterop|Diagnostics")]`
- `[Trait("Feature", "<feature-name>")]` - See valid values below
- `[Trait("RequiresExcel", "true")]` - All integration tests require Excel
- `[Trait("RunType", "OnDemand")]` - For session/lifecycle tests and diagnostic tests (slow, run only when explicitly requested)

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

**⚠️ CRITICAL: Always specify the test project explicitly to avoid running all test projects!**

### Core.Tests (Business Logic)
```bash
# Development (fast - excludes VBA)
dotnet test tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Diagnostic tests (validate patterns, slow ~20s each)
dotnet test tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj --filter "RunType=OnDemand&Layer=Diagnostics"

# VBA tests (manual only - requires VBA trust)
dotnet test tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"

# Specific feature
dotnet test tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj --filter "Feature=PowerQuery"
```

### ComInterop.Tests (Session/Batch Infrastructure)
```bash
# Session/batch changes (MANDATORY - see CRITICAL-RULES.md Rule 3)
dotnet test tests/ExcelMcp.ComInterop.Tests/ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"
```

### McpServer.Tests (End-to-End Tool Tests)
```bash
# All MCP tool tests
dotnet test tests/ExcelMcp.McpServer.Tests/ExcelMcp.McpServer.Tests.csproj

# Specific tool
dotnet test tests/ExcelMcp.McpServer.Tests/ExcelMcp.McpServer.Tests.csproj --filter "FullyQualifiedName~PowerQueryTool"
```

### CLI.Tests (Command-Line Interface)
```bash
# All CLI tests
dotnet test tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj

# Specific command
dotnet test tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj --filter "FullyQualifiedName~PowerQuery"
```

### Run Specific Test by Name
```bash
# Use full project path + filter
dotnet test tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj --filter "FullyQualifiedName~TestMethodName"
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

### Content Replacement Validation (CRITICAL)

**For operations that replace content (Update, Set, etc.), ALWAYS verify content was replaced, not merged/appended:**

```csharp
// ❌ WRONG: Only checks operation completed
var updateResult = await _commands.UpdateAsync(batch, queryName, newFile);
Assert.True(updateResult.Success);  // Doesn't prove content was replaced!

// ✅ CORRECT: Verify content was replaced, not merged
var updateResult = await _commands.UpdateAsync(batch, queryName, newFile);
Assert.True(updateResult.Success);

var viewResult = await _commands.ViewAsync(batch, queryName);
Assert.Equal(expectedContent, viewResult.Content);  // ✅ Content matches expected
Assert.DoesNotContain("OldContent", viewResult.Content);  // ✅ Old content gone!

// ✅ EVEN BETTER: Test multiple sequential updates (exposes merging bugs)
await _commands.UpdateAsync(batch, queryName, file1);
await _commands.UpdateAsync(batch, queryName, file2);
var viewResult = await _commands.ViewAsync(batch, queryName);
Assert.Equal(file2Content, viewResult.Content);  // ✅ Only file2 content present
Assert.DoesNotContain(file1Content, viewResult.Content);  // ✅ file1 content gone!
```

**Why Critical:** Bug report showed that UpdateAsync was **merging** M code instead of replacing it. Tests passed because they only checked `Success = true`, not actual content. The bug compounded with each update, corrupting queries progressively worse.

**Lesson:** "Operation completed" ≠ "Operation did the right thing". Always verify the actual result.

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| Shared test file | Each test creates unique file |
| Only test success flag | Verify actual Excel state |
| Save before assertions | Remove Save entirely |
| Save in middle of test | Only at end or in persistence test |
| Manual IDisposable | Use `IClassFixture<TempDirectoryFixture>` |
| .xlsx for VBA tests | Use `.xlsm` |
| "Accept both" assertions | Binary assertions only |
| Missing Feature trait | Add from valid feature list above |

## When Tests Fail

1. Run individually: `--filter "FullyQualifiedName=Namespace.Class.Method"`
2. Check file isolation (unique files?)
3. Check assertions (binary, not conditional?)
4. Check Save (removed unless persistence test?)
5. Verify Excel state (not just success flag?)

**Full checklist**: See CRITICAL-RULES.md Rule 12
