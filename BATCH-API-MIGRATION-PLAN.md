# Batch API Migration - Continuation Plan

## Current Status (as of 2025-10-27)

### ✅ Completed Work
- **Build Status:** ✅ 0 errors, 0 warnings
- **Tests Passing:** ✅ 86/86 (100%) - 61 CLI + 25 Core
- **PowerQueryCommandsTests:** ✅ Fully converted (30 test methods)
- **CLI Exception Handling:** ✅ All BeginBatchAsync calls wrapped in try-catch
- **Simple Tests Created:** ✅ ConnectionCommandsSimpleTests, SetupCommandsSimpleTests

### 📋 Test Files Currently Excluded

#### Core.Tests (.csproj exclusions)
```xml
<!-- Integration/Commands -->
<Compile Remove="Integration\Commands\CellCommandsTests.cs" />
<Compile Remove="Integration\Commands\ParameterCommandsTests.cs" />
<Compile Remove="Integration\Commands\FileCommandsTests.cs" />
<Compile Remove="Integration\Commands\DataModelTomCommandsTests.cs" />
<Compile Remove="Integration\Commands\DataModelCommandsTests.cs" />
<Compile Remove="Integration\Commands\ScriptCommandsTests.cs" />
<Compile Remove="Integration\Commands\CoreConnectionCommandsTests.cs" />
<Compile Remove="Integration\Commands\CoreConnectionCommandsExtendedTests.cs" />
<Compile Remove="Integration\Commands\VbaTrustDetectionTests.cs" />
<Compile Remove="Integration\Commands\SheetCommandsTests.cs" />
<Compile Remove="Integration\Commands\SetupCommandsTests.cs" />
<Compile Remove="Integration\Commands\PowerQueryWorkflowGuidanceTests.cs" />
<Compile Remove="Integration\Commands\PowerQueryPrivacyLevelTests.cs" />

<!-- Entire directories -->
<Compile Remove="RoundTrip\**" />
<Compile Remove="Unit\**" />

<!-- Helpers -->
<Compile Remove="Helpers\ConnectionTestHelper.cs" />
<Compile Remove="Helpers\DataModelTestHelper.cs" />
```

#### CLI.Tests (.csproj exclusions)
```xml
<Compile Remove="Integration\Commands\DataModelCommandsTests.cs" />
```

#### McpServer.Tests (.csproj exclusions)
```xml
<Compile Remove="**\*.cs" />  <!-- All tests excluded -->
```

---

## Plan A: Create More Simple Tests (Fast Track)

### Objective
Rapidly create minimal test coverage for all command types using the batch API pattern.

### Estimated Time
2-3 hours for all command types

### Test Creation Pattern
```csharp
public class XxxCommandsSimpleTests : IDisposable
{
    private readonly string _testFile;
    private readonly XxxCommands _commands;
    
    public XxxCommandsSimpleTests()
    {
        _testFile = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        _commands = new XxxCommands();
        
        // Create test file
        var fileCommands = new FileCommands();
        fileCommands.CreateEmptyAsync(_testFile).GetAwaiter().GetResult();
    }
    
    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Speed", "Fast")]
    public async Task BasicOperation_Success()
    {
        // Arrange
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        
        // Act
        var result = await _commands.SomeMethodAsync(batch, args);
        
        // Assert
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
    }
    
    public void Dispose()
    {
        if (File.Exists(_testFile)) File.Delete(_testFile);
    }
}
```

### Files to Create (Priority Order)

#### High Priority (Core Commands)
1. **CellCommandsSimpleTests.cs**
   - `GetValue_EmptyCell_ReturnsNull()`
   - `SetValue_ValidCell_Success()`
   - Location: `tests/ExcelMcp.Core.Tests/Integration/Commands/`
   - Estimated: 15 min

2. **ParameterCommandsSimpleTests.cs**
   - `List_EmptyWorkbook_ReturnsSuccess()`
   - `Create_NewParameter_Success()`
   - Location: `tests/ExcelMcp.Core.Tests/Integration/Commands/`
   - Estimated: 15 min

3. **FileCommandsSimpleTests.cs**
   - `CreateEmpty_ValidPath_CreatesFile()`
   - Location: `tests/ExcelMcp.Core.Tests/Integration/Commands/`
   - Estimated: 10 min

4. **SheetCommandsSimpleTests.cs**
   - `List_NewWorkbook_ReturnsDefaultSheet()`
   - `Create_NewSheet_Success()`
   - Location: `tests/ExcelMcp.Core.Tests/Integration/Commands/`
   - Estimated: 15 min

5. **ScriptCommandsSimpleTests.cs** (requires .xlsm)
   - `List_EmptyMacroFile_ReturnsSuccess()`
   - Location: `tests/ExcelMcp.Core.Tests/Integration/Commands/`
   - Estimated: 20 min

#### Medium Priority (Data Model)
6. **DataModelCommandsSimpleTests.cs**
   - `ListTables_EmptyDataModel_ReturnsSuccess()`
   - Location: `tests/ExcelMcp.Core.Tests/Integration/Commands/`
   - Estimated: 20 min

7. **DataModelTomCommandsSimpleTests.cs**
   - `GetModelInfo_EmptyModel_ReturnsSuccess()`
   - Location: `tests/ExcelMcp.Core.Tests/Integration/Commands/`
   - Estimated: 20 min

#### Lower Priority (Specialized)
8. **VbaTrustSimpleTests.cs**
   - `CheckVbaTrust_ReturnsResult()`
   - Location: `tests/ExcelMcp.Core.Tests/Integration/Commands/`
   - Estimated: 10 min

9. **PowerQueryWorkflowSimpleTests.cs**
   - `Import_BasicQuery_ShowsWorkflowHints()`
   - Location: `tests/ExcelMcp.Core.Tests/Integration/Commands/`
   - Estimated: 15 min

### Implementation Steps

1. **Create test file** using pattern above
2. **Build** to verify no compilation errors
3. **Run test** to verify it passes: `dotnet test --filter "FullyQualifiedName~SimpleTests"`
4. **Commit** after each file: `git commit -m "test: Add XxxCommandsSimpleTests"`
5. **Repeat** for next command type

### Success Criteria
- Each simple test file has 1-3 tests
- All tests use batch API pattern (`BeginBatchAsync`)
- All tests pass
- Build remains clean (0 errors, 0 warnings)

---

## Plan B: Re-enable and Convert Excluded Tests (Comprehensive)

### Objective
Systematically convert all excluded test files to use the batch API pattern.

### Estimated Time
8-12 hours for all files (can be split into multiple sessions)

### Conversion Pattern

#### Old Pattern (Removed)
```csharp
[Fact]
public void TestMethod()
{
    var result = ExcelSession.Execute(filePath, save: false, (excel, workbook) =>
    {
        // Use workbook directly
        dynamic sheets = workbook.Worksheets;
        return 0;
    });
}
```

#### New Pattern (Batch API)
```csharp
[Fact]
public async Task TestMethodAsync()
{
    await using var batch = await ExcelSession.BeginBatchAsync(filePath);
    var result = await _commands.MethodAsync(batch, args);
    
    Assert.True(result.Success);
}
```

### Files to Convert (Priority Order)

#### Phase 1: Integration Commands (High Value)
1. **FileCommandsTests.cs**
   - Methods: ~10 tests
   - Pattern: Convert to `FileCommands.CreateEmptyAsync()`, etc.
   - Estimated: 45 min

2. **SheetCommandsTests.cs**
   - Methods: ~20 tests
   - Pattern: Batch API for Read/Write/Create/Delete
   - Estimated: 90 min

3. **CellCommandsTests.cs**
   - Methods: ~15 tests
   - Pattern: GetValueAsync/SetValueAsync with batch
   - Estimated: 60 min

4. **ParameterCommandsTests.cs**
   - Methods: ~15 tests
   - Pattern: ListAsync/GetAsync/SetAsync with batch
   - Estimated: 60 min

5. **ScriptCommandsTests.cs**
   - Methods: ~18 tests
   - Pattern: ListAsync/ExportAsync/ImportAsync with batch
   - Estimated: 75 min

6. **PowerQueryWorkflowGuidanceTests.cs**
   - Methods: ~8 tests
   - Pattern: Import/Update with workflow hints
   - Estimated: 40 min

7. **PowerQueryPrivacyLevelTests.cs**
   - Methods: ~10 tests
   - Pattern: Import with privacy levels
   - Estimated: 50 min

#### Phase 2: Data Model Tests
8. **DataModelCommandsTests.cs** (Core)
   - Methods: ~15 tests
   - Pattern: ListTablesAsync/ListMeasuresAsync with batch
   - Estimated: 75 min

9. **DataModelTomCommandsTests.cs**
   - Methods: ~12 tests
   - Pattern: TOM API operations with batch
   - Estimated: 60 min

10. **CLI DataModelCommandsTests.cs**
    - Methods: ~10 tests
    - Pattern: CLI wrapper tests
    - Estimated: 50 min

#### Phase 3: Connections
11. **CoreConnectionCommandsTests.cs**
    - Methods: ~20 tests
    - Pattern: ListAsync/ViewAsync/ImportAsync with batch
    - Estimated: 90 min

12. **CoreConnectionCommandsExtendedTests.cs**
    - Methods: ~15 tests
    - Pattern: Advanced connection operations
    - Estimated: 75 min

#### Phase 4: Setup & Detection
13. **SetupCommandsTests.cs**
    - Methods: ~8 tests
    - Pattern: VBA trust detection
    - Estimated: 40 min

14. **VbaTrustDetectionTests.cs**
    - Methods: ~10 tests
    - Pattern: Trust state detection
    - Estimated: 50 min

#### Phase 5: Test Helpers (Required First)
15. **ConnectionTestHelper.cs**
    - Methods: 5 helper methods
    - Pattern: Convert Execute() to ExecuteAsync<int>()
    - **MUST DO FIRST** before connection tests
    - Estimated: 30 min

16. **DataModelTestHelper.cs**
    - Methods: 2 helper methods
    - Pattern: Convert Execute() to ExecuteAsync<int>()
    - **MUST DO FIRST** before data model tests
    - Estimated: 20 min

#### Phase 6: Unit Tests (Large Volume)
17. **Unit/\*\*/\*.cs** (Multiple files)
    - Count: ~50+ test files
    - Pattern: Various unit tests
    - Estimated: 4-6 hours
    - **LOW PRIORITY** - Can be done incrementally

#### Phase 7: RoundTrip Tests (Complex Workflows)
18. **RoundTrip/\*\*/\*.cs** (Multiple files)
    - Count: ~10 test files
    - Pattern: End-to-end workflows
    - Estimated: 2-3 hours
    - **MEDIUM PRIORITY** - Valuable but complex

#### Phase 8: MCP Server Tests
19. **All MCP Server tests** (Currently all excluded)
    - Count: ~30+ test files
    - Pattern: Tool invocation tests
    - Different challenge: Task<string> → ReadOnlySequence<byte>
    - Estimated: 3-4 hours
    - **SEPARATE EFFORT** - Requires different conversion pattern

### Conversion Workflow (Per File)

1. **Remove exclusion** from .csproj
2. **Build** and capture errors: `dotnet build 2>&1 | Select-String "error CS" > errors.txt`
3. **Analyze errors** to identify patterns
4. **Convert systematically:**
   - Update method signatures to `async Task`
   - Replace `ExcelSession.Execute()` with `BeginBatchAsync()`
   - Update Core command calls to Async versions
   - Update helper calls to Async versions
   - Fix assertions for async code
5. **Build** to verify: `dotnet build`
6. **Run tests** to verify: `dotnet test --filter "FullyQualifiedName~FileCommandsTests"`
7. **Commit** when all tests pass: `git commit -m "test: Convert FileCommandsTests to batch API"`
8. **Move to next file**

### Common Conversion Patterns

#### Pattern 1: ExcelSession.Execute() → ExecuteAsync()
```csharp
// OLD
var exitCode = ExcelSession.Execute(filePath, save: false, (excel, workbook) => {
    dynamic sheets = workbook.Worksheets;
    return 0;
});

// NEW
var exitCode = await ExcelSession.ExecuteAsync<int>(filePath, save: false, 
    async (context, ct) => {
        dynamic sheets = context.Book.Worksheets;
        return ValueTask.FromResult(0);
    });
```

#### Pattern 2: Core Commands with Batch
```csharp
// OLD
var result = _commands.Import(filePath, queryName, mCodeFile);

// NEW
await using var batch = await ExcelSession.BeginBatchAsync(filePath);
var result = await _commands.ImportAsync(batch, queryName, mCodeFile);
await batch.SaveAsync();
```

#### Pattern 3: Test Helpers
```csharp
// OLD
ConnectionTestHelper.CreateOdbcConnection(filePath);

// NEW
await ConnectionTestHelper.CreateOdbcConnectionAsync(filePath);
// or
ConnectionTestHelper.CreateOdbcConnectionAsync(filePath).GetAwaiter().GetResult();
```

#### Pattern 4: Assert.Throws → Assert.ThrowsAsync
```csharp
// OLD
Assert.Throws<ArgumentException>(() => _commands.Import(filePath, null));

// NEW
await Assert.ThrowsAsync<ArgumentException>(
    async () => await _commands.ImportAsync(batch, null));
```

### Success Criteria
- All excluded test files re-enabled
- All tests converted to batch API
- Build clean (0 errors, 0 warnings)
- All tests passing
- No use of removed `Execute()` method

---

## Hybrid Approach (Recommended)

### Phase 1: Quick Coverage (Plan A - 2 hours)
Create simple tests for all command types to ensure basic coverage.

### Phase 2: High-Value Conversions (Plan B - 4 hours)
Convert the most important test files:
1. FileCommandsTests
2. SheetCommandsTests
3. CellCommandsTests
4. ParameterCommandsTests
5. ScriptCommandsTests

### Phase 3: Incremental Conversion (Plan B - ongoing)
Convert remaining tests in future PRs as needed.

---

## Commands for GitHub Coding Agent

### To Start Plan A (Create Simple Tests)
```bash
# 1. Create first simple test file
# Use the pattern from ConnectionCommandsSimpleTests.cs
# Start with: CellCommandsSimpleTests.cs

# 2. Build and test
dotnet build
dotnet test --filter "FullyQualifiedName~SimpleTests"

# 3. Commit
git add tests/ExcelMcp.Core.Tests/Integration/Commands/CellCommandsSimpleTests.cs
git commit -m "test: Add CellCommandsSimpleTests with batch API"

# Repeat for each command type
```

### To Start Plan B (Convert Existing Tests)
```bash
# 1. Re-enable one test file in .csproj
# Edit: tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj
# Remove line: <Compile Remove="Integration\Commands\FileCommandsTests.cs" />

# 2. Build and capture errors
dotnet build 2>&1 | Select-String "error CS" | Out-File -FilePath conversion-errors.txt

# 3. Analyze and convert
# Use patterns documented above

# 4. Verify
dotnet build
dotnet test --filter "FullyQualifiedName~FileCommandsTests"

# 5. Commit
git add tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj
git add tests/ExcelMcp.Core.Tests/Integration/Commands/FileCommandsTests.cs
git commit -m "test: Convert FileCommandsTests to batch API (XX tests)"
```

### Progress Tracking
```bash
# Check enabled test count
dotnet test --list-tests --no-build 2>&1 | Select-String "^\s+Sbroenne" | Measure-Object

# Run all tests
dotnet test

# Check build status
dotnet build -c Debug
```

---

## Key Reference Files

### Working Examples (Use as Templates)
- `tests/ExcelMcp.Core.Tests/Integration/Commands/PowerQueryCommandsTests.cs` - ✅ Fully converted (30 tests)
- `tests/ExcelMcp.Core.Tests/Integration/Commands/ConnectionCommandsSimpleTests.cs` - ✅ Simple test pattern
- `tests/ExcelMcp.Core.Tests/Integration/Commands/SetupCommandsSimpleTests.cs` - ✅ Simple test pattern

### Core API Reference
- `src/ExcelMcp.Core/Session/ExcelSession.cs` - Batch API implementation
- `src/ExcelMcp.Core/Commands/*.cs` - All command implementations (Async methods)
- `src/ExcelMcp.Core/Models/ResultTypes.cs` - Result type definitions

---

## Notes for Coding Agent

1. **Always build before running tests** to catch compilation errors early
2. **Commit frequently** - after each successful file conversion
3. **Use parallel test execution** for faster feedback: `dotnet test --parallel`
4. **Check test categories**: Integration tests require Excel installed
5. **OnDemand tests** (pool cleanup) should remain excluded - they're run manually
6. **Simple tests are faster** than converting complex test files
7. **Start with helpers** if converting connection/data model tests
8. **MCP Server tests** need different conversion pattern (not covered in this plan)

---

## Estimated Total Time

### Plan A Only (Simple Tests)
- **2-3 hours** for complete coverage of all command types
- **Result:** Basic smoke tests for all functionality

### Plan B Only (Full Conversion)
- **8-12 hours** for all integration tests
- **15-20 hours** including unit tests and round-trip tests
- **Result:** Comprehensive test coverage

### Hybrid Approach (Recommended)
- **Phase 1:** 2 hours (simple tests)
- **Phase 2:** 4 hours (high-value conversions)
- **Phase 3:** Ongoing (remaining tests)
- **Total:** 6 hours for good coverage, expand as needed

---

## Success Metrics

### Minimum Viable (Current State)
- ✅ Build: 0 errors
- ✅ Tests: 86 passing
- ✅ Coverage: Core PowerQuery functionality

### Plan A Complete
- ✅ Build: 0 errors
- ✅ Tests: ~100+ passing
- ✅ Coverage: Smoke tests for all command types

### Plan B Complete (Phases 1-4)
- ✅ Build: 0 errors
- ✅ Tests: ~200+ passing
- ✅ Coverage: Comprehensive integration tests

### Full Migration Complete
- ✅ Build: 0 errors
- ✅ Tests: 300+ passing
- ✅ Coverage: All tests converted
- ✅ No excluded test files
