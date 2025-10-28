---
applyTo: "tests/**/*.cs"
---

# Testing Strategy

> **Comprehensive guide to ExcelMcp's three-tier testing approach**

## Test Architecture

```
tests/
├── ExcelMcp.Core.Tests/
│   ├── Unit/           # Fast, no Excel (2-5 sec)
│   ├── Integration/    # Medium, requires Excel (1-15 min)
│   └── RoundTrip/      # Slow, complex workflows (3-10 min each)
├── ExcelMcp.McpServer.Tests/
│   ├── Unit/
│   ├── Integration/
│   └── RoundTrip/
└── ExcelMcp.CLI.Tests/
    ├── Unit/
    └── Integration/
```

---

## Test Traits (REQUIRED)

```csharp
// Unit Tests
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core|CLI|McpServer")]

// Integration Tests  
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "PowerQuery|VBA|Worksheets|Files")]
[Trait("RequiresExcel", "true")]

// Round Trip Tests
[Trait("Category", "RoundTrip")]
[Trait("Speed", "Slow")]
[Trait("Feature", "EndToEnd|MCPProtocol|Workflows")]
[Trait("RequiresExcel", "true")]

// OnDemand Tests (pool cleanup, stress tests)
[Trait("RunType", "OnDemand")]
[Trait("Speed", "Slow")]
```

---

## Development Workflow

```bash
# Daily development (fast feedback)
dotnet test --filter "Category=Unit&RunType!=OnDemand"

# Pre-commit validation
dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand"

# Pool code changes (MANDATORY)
dotnet test --filter "RunType=OnDemand" --list-tests  # Verify 5 tests
dotnet test --filter "RunType=OnDemand"               # Run (3-5 min)

# Full validation
dotnet test --filter "Category=RoundTrip"
```

---

## OnDemand Test Strategy

### Why OnDemand Tests Exist

**Problem:** Pool cleanup tests verify Excel.exe process termination via COM interop, which requires Excel installed. GitHub Actions doesn't have Excel.

**Solution:** Mark with `[Trait("RunType", "OnDemand")]` for local-only execution.

### Two-Tier Approach

1. **Unit Tests** (CI/CD):
   - Verify pool logic, semaphore behavior, capacity enforcement
   - No Excel required
   - Fast (2-5 seconds)

2. **OnDemand Tests** (Local):
   - Verify actual COM cleanup and process termination
   - Requires Excel installed
   - Takes 3-5 minutes
   - **MANDATORY before committing pool changes**

### When to Run OnDemand Tests

✅ **ALWAYS:**
- Modifying `ExcelInstancePool.cs`
- Modifying `ExcelHelper.cs` pooling code
- Changing semaphore logic
- Before releasing with pool changes

✅ **Optional:**
- Weekly regression testing
- After .NET/Excel upgrades

❌ **Never:**
- CI/CD pipelines (no Excel)
- Quick development iterations

### How to Run

```bash
# STEP 1: Verify filter
dotnet test --filter "RunType=OnDemand" --list-tests --nologo

# STEP 2: Close all Excel instances

# STEP 3: Run tests (3-5 minutes)
dotnet test --filter "RunType=OnDemand" --nologo

# STEP 4: ALL 5 must pass before commit
```

### What OnDemand Tests Verify

1. **Semaphore race prevention** - No TOCTOU bugs, capacity never exceeded
2. **COM cleanup** - Excel.exe processes terminate after disposal
3. **Eviction behavior** - Removed instances clean up immediately
4. **Stress resilience** - 50+ parallel operations don't leak processes
5. **Fixture disposal** - Test cleanup disposes all instances

### Tests with Side Effects

**Any test that spawns/terminates Excel processes should be marked OnDemand:**

```csharp
[Trait("Category", "Unit")]
[Trait("RunType", "OnDemand")]  // At class level for documentation
public class StaThreadingTests
{
    [Fact]
    [Trait("RunType", "OnDemand")]  // At method level for filtering
    public async Task ExecuteAsync_WithStaThreading_NoProcessLeak()
    {
        // Test that verifies Excel.exe cleanup
    }
}
```

**Why:** These tests have side effects that can interfere with parallel test execution and require Excel to be installed.

**Examples:**
- `StaThreadingTests` - Verifies Excel process cleanup
- Pool cleanup tests - Terminates Excel instances
- Any test that counts running Excel processes

---

## Test Naming Standards

### Layer Prefixes (REQUIRED)

```csharp
// CLI Tests
public class CliFileCommandsTests { }
public class CliPowerQueryCommandsTests { }

// Core Tests
public class CoreFileCommandsTests { }
public class CorePowerQueryCommandsTests { }

// MCP Server Tests
public class McpServerRoundTripTests { }
public class ExcelMcpServerTests { }
```

**Why:** Prevents FQDN conflicts and enables precise test filtering.

---

## Test Brittleness Prevention

### Common Issues

1. **Shared State**
```csharp
// ❌ BAD
private readonly string _testFile = "shared.xlsx";

// ✅ GOOD
string testFile = $"test-{Guid.NewGuid():N}.xlsx";
```

2. **Invalid Assumptions**
```csharp
// ❌ BAD - Assumes empty cell has value
Assert.NotNull(result.Value);

// ✅ GOOD - Tests realistic Excel behavior
Assert.True(result.Success);
Assert.Null(result.ErrorMessage);
```

3. **Type Mismatches**
```csharp
// ❌ BAD - String vs numeric comparison
Assert.Equal("30", result.Value);

// ✅ GOOD - Convert to string
Assert.Equal("30", result.Value?.ToString());
```

---

## CI/CD Strategy

```yaml
# GitHub Actions (no Excel)
jobs:
  unit-tests:
    steps:
    - run: dotnet test --filter "Category=Unit&RunType!=OnDemand"
      # ✅ Fast, no Excel required
  
  integration-tests:
    steps:
    - run: dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand"
      # ⚠️ Skips OnDemand pool tests (no Excel)
```

---

## Test Filter Validation

**⚠️ ALWAYS verify filters before running:**

```bash
# Verify what will run
dotnet test --filter "FullyQualifiedName~ExcelPoolCleanupTests" --list-tests

# Check count and names match expectations

# Then run
dotnet test --filter "FullyQualifiedName~ExcelPoolCleanupTests"
```

---

## Performance Targets

- **Unit**: ~46 tests, 2-5 seconds
- **Integration**: ~91+ tests, 13-15 minutes  
- **RoundTrip**: ~10+ tests, 3-10 minutes each
- **OnDemand**: 5 tests, 3-5 minutes
- **Total**: 150+ tests across all layers

---

## Test Layer Separation of Concerns (CRITICAL)

### ⚠️ PROBLEM: Test Duplication Across Layers

**Current Issue** (as of 2025-10-27):
- Core: 181 tests ✅ CORRECT
- CLI: 63 tests ❌ **51+ duplicating Core business logic**
- MCP: 3 tests ⚠️ INSUFFICIENT

### Test Responsibility Matrix

| Concern | Core | CLI | MCP Server |
|---------|------|-----|------------|
| **Excel COM Operations** | ✅ | ❌ | ❌ |
| **Power Query M Code** | ✅ | ❌ | ❌ |
| **VBA Operations** | ✅ | ❌ | ❌ |
| **Worksheet Operations** | ✅ | ❌ | ❌ |
| **File Creation** | ✅ | ❌ | ❌ |
| **Result Object Validation** | ✅ | ❌ | ❌ |
| **Error Handling Logic** | ✅ | ❌ | ❌ |
| **Argument Parsing** | ❌ | ✅ | ❌ |
| **Exit Code Mapping** | ❌ | ✅ | ❌ |
| **Console Formatting** | ❌ | ✅ | ❌ |
| **User Prompts** | ❌ | ✅ | ❌ |
| **JSON Serialization** | ❌ | ❌ | ✅ |
| **MCP Protocol** | ❌ | ❌ | ✅ |
| **Tool Action Routing** | ❌ | ❌ | ✅ |
| **Parameter Binding** | ❌ | ❌ | ✅ |

### Core Tests - ✅ CORRECT Approach

```csharp
/// <summary>
/// Core tests - Verify business logic and Excel COM operations
/// DO NOT test console output or argument parsing
/// </summary>
[Trait("Category", "Integration")]
[Trait("Layer", "Core")]
public class CoreFileCommandsTests : IDisposable
{
    private readonly FileCommands _fileCommands;
    
    [Fact]
    public async Task CreateEmpty_WithValidPath_ReturnsSuccessResult()
    {
        string testFile = Path.Combine(_tempDir, "TestFile.xlsx");
        
        // ✅ Tests Core business logic
        var result = await _fileCommands.CreateEmptyAsync(testFile);
        
        // ✅ Validates Result object
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.Equal("create-empty", result.Action);
        Assert.True(File.Exists(testFile));
    }
}
```

**What Core tests SHOULD verify**:
- Excel COM interop operations
- Power Query M code operations
- VBA operations
- Worksheet/cell operations
- Result object properties
- Error handling logic
- Business rule validation

### CLI Tests - ❌ AVOID Duplicating Core Logic

```csharp
/// <summary>
/// CLI tests - Verify argument parsing and exit codes ONLY
/// DO NOT test Excel operations (those are in Core)
/// </summary>
[Trait("Category", "Unit")]
[Trait("Layer", "CLI")]
public class CliArgumentParsingTests
{
    // ✅ CORRECT - Tests CLI-specific concern
    [Theory]
    [InlineData(new string[] { "create-empty" }, 1)] // Missing file
    [InlineData(new string[] { "create-empty", "test.txt" }, 1)] // Invalid extension
    public void CreateEmpty_InvalidArgs_ReturnsOne(string[] args, int expected)
    {
        var commands = new FileCommands();
        int exitCode = commands.CreateEmpty(args);
        Assert.Equal(expected, exitCode);
    }
    
    // ❌ WRONG - Tests Core business logic
    [Fact]
    public void CreateEmpty_WithValidPath_ReturnsZeroAndCreatesFile()
    {
        int exitCode = _cliCommands.CreateEmpty(args);
        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(testFile)); // ❌ Core concern!
    }
}
```

**What CLI tests SHOULD verify**:
- Argument parsing (missing args, invalid format)
- Exit code mapping (0 for success, 1 for error)
- Console output formatting (optional)
- User prompts/confirmations

**What CLI tests should NOT verify**:
- File creation/deletion
- Excel operations
- Power Query operations
- Data manipulation
- Business logic (all in Core)

### MCP Server Tests - ⚠️ Need More Coverage

```csharp
/// <summary>
/// MCP tests - Verify JSON serialization and MCP protocol ONLY
/// DO NOT test Excel operations (those are in Core)
/// </summary>
[Trait("Category", "Unit")]
[Trait("Layer", "McpServer")]
public class ExcelPowerQueryToolTests
{
    [Fact]
    public void ExcelPowerQuery_ListAction_SerializesCorrectly()
    {
        // ✅ Test JSON serialization of Core results
        var coreResult = new PowerQueryListResult { ... };
        string json = ExcelTools.ExcelPowerQuery("list", ...);
        
        // Verify JSON structure
        var deserialized = JsonSerializer.Deserialize<...>(json);
        Assert.Equal(expected, deserialized);
    }
    
    [Fact]
    public void ExcelPowerQuery_ErrorResult_SerializesToMcpError()
    {
        // ✅ Test error serialization
    }
}
```

**What MCP tests SHOULD verify**:
- JSON serialization of Result objects
- MCP protocol compliance
- Tool action routing
- Parameter binding/validation
- Error serialization
- JSON schema validation

**What MCP tests should NOT verify**:
- Excel operations
- Power Query logic
- File creation
- Business logic (all in Core)

---

## Resource Management in Tests (CRITICAL)

### ❌ PROBLEM: Tests Handling GC/Dispose

**WRONG** - Tests should NOT manually trigger GC:
```csharp
// In test Dispose()
public void Dispose()
{
    foreach (string file in _createdFiles)
    {
        try { File.Delete(file); }
        catch (IOException)
        {
            System.Threading.Thread.Sleep(1000);
            GC.Collect();  // ❌ Tests shouldn't do this!
            GC.WaitForPendingFinalizers();  // ❌ Core design issue!
        }
    }
}
```

**Why this is wrong**:
- Tests are working around Core resource management issues
- If tests need manual GC, Core isn't properly cleaning up COM objects
- GC timing is non-deterministic (unreliable tests)
- Violates separation of concerns

**CORRECT** - Core should handle all COM cleanup:
```csharp
// In ExcelSession/ExcelBatch
public async ValueTask DisposeAsync()
{
    // Proper COM cleanup
    if (_book != null)
    {
        ComHelper.ReleaseComObject(ref _book);
    }
    if (_app != null)
    {
        _app.Quit();
        ComHelper.ReleaseComObject(ref _app);
    }
    
    // Force COM cleanup
    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();
}
```

**Then tests become simple**:
```csharp
// Test Dispose() - just cleanup test artifacts
public void Dispose()
{
    try
    {
        if (Directory.Exists(_tempDir))
        {
            Directory.Delete(_tempDir, recursive: true);
        }
    }
    catch { /* Cleanup failure is non-critical */ }
}
```

**Exception: StaThreadingTests.cs** (CORRECT usage):
```csharp
[Fact]
public async Task BeginBatchAsync_DisposesCorrectly_NoExcelProcessLeak()
{
    // ... test operations ...
    
    await Task.Delay(5000);
    GC.Collect();  // ✅ CORRECT - Explicitly testing COM cleanup
    GC.WaitForPendingFinalizers();
    GC.Collect();
    
    var endingProcesses = Process.GetProcessesByName("EXCEL");
    Assert.True(endingCount <= startingCount, "Excel process leak!");
}
```

This is **CORRECT** because:
- Test is explicitly verifying COM cleanup behavior
- GC calls are part of the test assertion
- Documented in test name and comments

---

## Key Lessons

1. **OnDemand pattern** - Essential for Excel-dependent tests that can't run in CI/CD
2. **Test isolation** - Save/restore global state (like `ExcelHelper.InstancePool`)
3. **Realistic data** - Use test helpers to create real Excel objects
4. **Layer prefixes** - Prevent FQDN conflicts in test class names
5. **Complete traits** - All tests MUST have Category, Speed, Layer traits
6. **Separation of concerns** - Core tests business logic, CLI tests argument parsing, MCP tests JSON serialization
7. **No test duplication** - If Core already tests it, don't duplicate in CLI/MCP
8. **Core handles cleanup** - Tests should NOT manually trigger GC (except when testing GC behavior)
