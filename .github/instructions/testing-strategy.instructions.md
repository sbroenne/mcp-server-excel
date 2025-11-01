---
applyTo: "tests/**/*.cs"
---

# Testing Strategy

> **Three-tier testing: Unit (fast) → Integration (Excel) → OnDemand (session cleanup)**

## 📋 Test Class Compliance Checklist

**Every new test class MUST follow these rules to prevent systematic issues:**

### ✅ File Organization
- [ ] Test class uses `partial` keyword for multi-file organization (if needed)
- [ ] File name matches class name (e.g., `ConnectionCommandsTests.cs` contains `ConnectionCommandsTests`)
- [ ] Files organized by feature in subdirectories (e.g., `Commands/Connection/`)
- [ ] Related partial files use descriptive suffixes (e.g., `ConnectionCommandsTests.List.cs`, `ConnectionCommandsTests.View.cs`)

### ✅ Test Fixture Setup
- [ ] Test class implements `IClassFixture<TempDirectoryFixture>` for integration tests
- [ ] Constructor accepts `TempDirectoryFixture` via dependency injection
- [ ] Store temp directory in `private readonly string _tempDir` field
- [ ] **NEVER** manually implement `IDisposable` for temp directory cleanup (fixture handles it)
- [ ] **NEVER** create temp directory in constructor (fixture provides it)

### ✅ Test File Isolation
- [ ] Each test creates its own unique file using `CoreTestHelper.CreateUniqueTestFileAsync()`
- [ ] **NEVER** share a single test file across multiple tests
- [ ] **NEVER** reuse file paths between tests
- [ ] Pass `_tempDir` (from fixture) to `CreateUniqueTestFileAsync()`
- [ ] Use test class name and test method name in file creation for traceability

### ✅ File Extension Requirements
- [ ] VBA tests MUST use `.xlsm` extension (macro-enabled workbooks)
- [ ] Standard tests use `.xlsx` extension (unless VBA required)
- [ ] CSV/data files use appropriate extensions (`.csv`, `.txt`)
- [ ] Pass extension parameter to `CoreTestHelper.CreateUniqueTestFileAsync()`
- [ ] **NEVER** rename files to change format (e.g., `.xlsx` → `.xlsm` fails)

### ✅ Test Assertions
- [ ] Use binary assertions: `Assert.True(result.Success, $"Reason: {result.ErrorMessage}")`
- [ ] **NEVER** use "accept both" patterns (if-success-pass, if-error-pass)
- [ ] Include descriptive failure messages in assertions
- [ ] Use `Skip` attribute if test requires unavailable features, not conditional returns

### ✅ Required Traits
- [ ] `[Trait("Category", "Integration")]` or `[Trait("Category", "Unit")]`
- [ ] `[Trait("Speed", "Medium")]` or `[Trait("Speed", "Fast")]`
- [ ] `[Trait("Layer", "Core|CLI|McpServer")]`
- [ ] `[Trait("RequiresExcel", "true")]` for integration tests
- [ ] `[Trait("Feature", "FeatureName")]` for grouping

### ✅ Helper Method Usage
- [ ] Use `CoreTestHelper.CreateUniqueTestFileAsync()` for ALL test file creation
- [ ] **NEVER** duplicate file creation logic in test classes
- [ ] For Excel files: Pass test class name, method name, temp dir, extension
- [ ] For data files: Pass test class name, method name, temp dir, extension, content
- [ ] **NEVER** create local helper methods that duplicate `CoreTestHelper` functionality

### ✅ Batch API Pattern
- [ ] All Core commands accept `IExcelBatch batch` as first parameter
- [ ] Tests create batch with `await ExcelSession.BeginBatchAsync(testFile)`
- [ ] Use `await using var batch` for automatic disposal
- [ ] Call `await batch.SaveAsync()` only if test modifies data

### ✅ Async/Await Patterns
- [ ] Test methods use `async Task` (not `async void`)
- [ ] All async operations properly awaited
- [ ] **NEVER** use `ValueTask.FromResult()` wrapper in `batch.ExecuteAsync()` (returns directly)
- [ ] Use `await using` for `IAsyncDisposable` resources

### ✅ Process Cleanup (Session Tests Only)
- [ ] Session/lifecycle tests implement `IAsyncLifetime`
- [ ] `InitializeAsync()` kills existing Excel processes before tests
- [ ] Add adequate delays for COM cleanup (5-7 seconds with forced GC)
- [ ] **ONLY** for tests that verify process lifecycle, not regular integration tests

## Quick Reference: Correct Test Class Template

```csharp
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Feature;

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
        // Arrange - Create unique test file
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(FeatureCommandsTests), 
            nameof(Operation_Scenario_ExpectedResult), 
            _tempDir,
            ".xlsx");  // or ".xlsm" for VBA tests

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _commands.OperationAsync(batch, args);

        // Assert
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        Assert.NotNull(result.Data);
    }
}
```

## Test Architecture

```
tests/
├── ExcelMcp.Core.Tests/      # Unit + Integration
├── ExcelMcp.McpServer.Tests/ # Unit + Integration
├── ExcelMcp.CLI.Tests/        # Unit + Integration
└── ExcelMcp.ComInterop.Tests/ # Unit + OnDemand
```

## Required Traits

```csharp
// Unit Tests (fast, no Excel)
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core|CLI|McpServer|ComInterop")]

// Integration Tests (requires Excel)
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("RequiresExcel", "true")]

// OnDemand Tests (session cleanup, stress tests)
[Trait("RunType", "OnDemand")]
[Trait("Speed", "Slow")]
```

## Development Workflow

```bash
# Fast feedback during development
dotnet test --filter "Category=Unit&RunType!=OnDemand"

# Pre-commit validation
dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand"

# Session/batch code changes (MANDATORY - see CRITICAL-RULES.md)
dotnet test --filter "RunType=OnDemand"
```

## ⚠️ CRITICAL: Test File Isolation

### ❌ ANTI-PATTERN: Shared Test File

```csharp
public class MyTests
{
    private readonly string _testExcelFile;  // ❌ WRONG: Single file for all tests
    
    public MyTests()
    {
        _testExcelFile = Path.Combine(_tempDir, "TestFile.xlsx");
        CreateTestFile(_testExcelFile);  // ❌ Created once, used by all tests
    }
    
    [Fact]
    public async Task Test1() 
    {
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);  // ❌ Shared file
        // ... test modifies file ...
    }
    
    [Fact]
    public async Task Test2()
    {
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);  // ❌ Same file!
        // ... test sees pollution from Test1 ...
    }
}
```

**Why This Is Catastrophic:**
- **Test Pollution**: Each test modifies shared file, contaminating subsequent tests
- **Order Dependency**: Tests pass/fail based on execution order
- **File Locks**: Excel processes holding file locks cause hanging
- **Hard to Debug**: File state contaminated by previous tests
- **Violates Isolation**: Tests are NOT independent

### ✅ CORRECT Pattern: Unique File Per Test

```csharp
public class MyTests
{
    private readonly string _tempDir;
    private readonly IFileCommands _fileCommands;
    
    public MyTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"MyTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _fileCommands = new FileCommands();
    }
    
    private async Task<string> CreateTestFileAsync(string fileName)
    {
        var filePath = Path.Combine(_tempDir, fileName);
        var result = await _fileCommands.CreateEmptyAsync(filePath);
        if (!result.Success) throw new InvalidOperationException($"Failed: {result.ErrorMessage}");
        
        // Add test data setup here
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        // ... setup code ...
        await batch.SaveAsync();
        
        return filePath;
    }
    
    [Fact]
    public async Task Test1()
    {
        // ✅ CORRECT: Each test gets its own file
        var testFile = await CreateTestFileAsync("Test1.xlsx");
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        // ... test is isolated ...
    }
    
    [Fact]
    public async Task Test2()
    {
        // ✅ CORRECT: Completely independent file
        var testFile = await CreateTestFileAsync("Test2.xlsx");
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        // ... no pollution from Test1 ...
    }
}
```

**Benefits:**
- ✅ **Complete Isolation**: Each test has pristine file
- ✅ **Order Independence**: Tests can run in any order
- ✅ **Parallel Safe**: Tests can run concurrently
- ✅ **Easy Debug**: File state is predictable
- ✅ **No File Locks**: Each file accessed by only one test

## ⚠️ CRITICAL: No "Accept Both" Tests

### ❌ CATASTROPHIC Pattern
```csharp
// Test always passes - feature can be 100% broken!
if (result.Success)
{
    Assert.True(result.Success);
}
else
{
    Assert.True(result.ErrorMessage.Contains("acceptable"));
}
```

### ✅ CORRECT Patterns

**Binary Assertion (Preferred):**
```csharp
Assert.True(result.Success, $"Must succeed: {result.ErrorMessage}");
```

**Skip if Unavailable:**
```csharp
if (!featureAvailable)
{
    _output.WriteLine("Skipping: Feature not available");
    return;
}
Assert.True(result.Success);
```

## OnDemand Tests

**Purpose:** Verify Excel.exe process cleanup (requires Excel, 3-5 min)

**When to run:**
- ✅ Modifying `ExcelSession.cs`, `ExcelBatch.cs`, or `ExcelHelper.cs`
- ❌ Never in CI/CD (no Excel)

```bash
dotnet test --filter "RunType=OnDemand" --list-tests
dotnet test --filter "RunType=OnDemand"
```

## Batch API Pattern

```csharp
// Core Commands
public async Task<OperationResult> MethodAsync(IExcelBatch batch, string arg)
{
    return await batch.ExecuteAsync(async (ctx, ct) =>
    {
        // Use ctx.Book for workbook operations
        return new OperationResult { Success = true };
    });
}

// Tests
[Fact]
public async Task TestMethod()
{
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    var result = await _commands.MethodAsync(batch, arg);
    Assert.True(result.Success);
}
```

## Layer Separation

| Concern | Core | CLI | MCP |
|---------|------|-----|-----|
| Excel COM | ✅ | ❌ | ❌ |
| Business Logic | ✅ | ❌ | ❌ |
| Argument Parsing | ❌ | ✅ | ❌ |
| Exit Codes | ❌ | ✅ | ❌ |
| JSON Protocol | ❌ | ❌ | ✅ |

**Rule:** Core tests business logic once. CLI tests parsing. MCP tests JSON.

## Test Naming

Use layer prefixes to avoid FQDN conflicts:

```csharp
public class CliFileCommandsTests { }      // CLI layer
public class CoreFileCommandsTests { }     // Core layer
public class McpServerRoundTripTests { }   // MCP layer
```

## Performance Targets

- **Unit**: ~46 tests, 2-5 sec
- **Integration**: ~91+ tests, 13-15 min
- **OnDemand**: 5 tests, 3-5 min
- **Total**: 150+ tests

## Key Principles

1. **Binary assertions** - Pass OR fail, never both
2. **OnDemand for side effects** - Excel process spawn/cleanup
3. **Layer prefixes** - Prevent naming conflicts
4. **Batch API** - All Core methods use `IExcelBatch`
5. **No duplication** - Core tests business logic once
6. **Unique file per test** - Each test creates its own isolated Excel file

## Common Mistakes & How to Avoid Them

### 1. ❌ Manual IDisposable Implementation
```csharp
// WRONG: Duplicate cleanup code in every test class
public class MyTests : IDisposable
{
    private readonly string _tempDir;
    
    public MyTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"MyTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }
    
    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }
}
```

**✅ CORRECT: Use TempDirectoryFixture**
```csharp
public class MyTests : IClassFixture<TempDirectoryFixture>
{
    private readonly string _tempDir;
    
    public MyTests(TempDirectoryFixture fixture)
    {
        _tempDir = fixture.TempDir;
    }
    // No Dispose needed - fixture handles cleanup!
}
```

### 2. ❌ Wrong File Extension for VBA Tests
```csharp
// WRONG: Creates .xlsx then tries to use VBA
var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
    nameof(MyTests), nameof(MyTest), _tempDir);  // Defaults to .xlsx
    
// Excel will reject VBA operations!
var result = await _scriptCommands.ImportAsync(batch, "Module1", vbaFile);
```

**✅ CORRECT: Explicitly specify .xlsm**
```csharp
var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
    nameof(MyTests), nameof(MyTest), _tempDir, ".xlsm");  // Macro-enabled
```

### 3. ❌ Using ValueTask.FromResult in batch.ExecuteAsync
```csharp
// WRONG: ExecuteAsync already wraps the return value
return await batch.ExecuteAsync(async (ctx, ct) =>
{
    var value = GetSomeValue();
    return ValueTask.FromResult(value);  // Double-wrapping!
});
```

**✅ CORRECT: Return directly**
```csharp
return await batch.ExecuteAsync(async (ctx, ct) =>
{
    var value = GetSomeValue();
    return value;  // ExecuteAsync handles wrapping
});
```

### 4. ❌ Duplicate File Creation Helpers
```csharp
// WRONG: Local helper duplicates CoreTestHelper functionality
private string CreateTestCsvFile(string testName)
{
    var fileName = $"{testName}_{Guid.NewGuid():N}.csv";
    var filePath = Path.Combine(_tempDir, fileName);
    File.WriteAllText(filePath, "Name,Value\nTest,123");
    return filePath;
}
```

**✅ CORRECT: Use CoreTestHelper**
```csharp
var csvFile = await CoreTestHelper.CreateUniqueTestFileAsync(
    nameof(MyTests), testName, _tempDir, ".csv", "Name,Value\nTest,123");
```

### 5. ❌ Forgetting Process Cleanup in Session Tests
```csharp
// WRONG: Baseline polluted by existing Excel processes
[Fact]
public async Task CreateNewAsync_NoExcelProcessLeak()
{
    var initialCount = Process.GetProcessesByName("EXCEL").Length;
    // Test may fail if Excel already running!
}
```

**✅ CORRECT: Implement IAsyncLifetime**
```csharp
public class ExcelSessionTests : IAsyncLifetime
{
    public async Task InitializeAsync()
    {
        // Kill existing Excel processes before each test
        foreach (var proc in Process.GetProcessesByName("EXCEL"))
        {
            proc.Kill();
            await proc.WaitForExitAsync();
        }
        GC.Collect();
        GC.WaitForPendingFinalizers();
        await Task.Delay(7000);  // Adequate cleanup time
    }
    
    public Task DisposeAsync() => Task.CompletedTask;
}
```

### 6. ❌ Shared Test Files
```csharp
// WRONG: All tests use same file
private readonly string _sharedFile;

public MyTests()
{
    _sharedFile = Path.Combine(_tempDir, "Shared.xlsx");
    CreateFile(_sharedFile);
}

[Fact] public async Task Test1() { /* uses _sharedFile */ }
[Fact] public async Task Test2() { /* uses _sharedFile - POLLUTED! */ }
```

**✅ CORRECT: Unique file per test**
```csharp
[Fact]
public async Task Test1()
{
    var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
        nameof(MyTests), nameof(Test1), _tempDir);
    // Fresh file, no pollution
}

[Fact]
public async Task Test2()
{
    var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
        nameof(MyTests), nameof(Test2), _tempDir);
    // Completely independent file
}
```

### 7. ❌ Missing Required Content for Data Files
```csharp
// WRONG: CSV file created without content
var csvFile = await CoreTestHelper.CreateUniqueTestFileAsync(
    nameof(MyTests), nameof(MyTest), _tempDir, ".csv");
// Throws ArgumentNullException!
```

**✅ CORRECT: Provide content parameter**
```csharp
var csvFile = await CoreTestHelper.CreateUniqueTestFileAsync(
    nameof(MyTests), nameof(MyTest), _tempDir, ".csv", 
    "Name,Value\nTest1,100\nTest2,200");
```


