---
applyTo: "tests/**/*.cs"
---

# Testing Strategy

> **Three-tier testing: Unit (fast) → Integration (Excel) → OnDemand (session cleanup)**

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

