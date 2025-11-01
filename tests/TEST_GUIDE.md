# ExcelMcp Testing Guide

> **Complete guide to running, writing, and organizing tests for ExcelMcp**

## Quick Start

```bash
# Development (fast feedback, no Excel required - excludes VBA)
dotnet test --filter "Category=Unit&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Pre-commit (comprehensive, requires Excel - excludes VBA)
dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Session/batch code changes (MANDATORY when modifying session/batch code)
dotnet test --filter "RunType=OnDemand"

# VBA tests only (manual, requires VBA trust enabled)
dotnet test --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"
```

---

## Test Architecture

### Three-Tier Testing Strategy

```
tests/
├── ExcelMcp.Core.Tests/      # Core business logic (Unit + Integration)
├── ExcelMcp.McpServer.Tests/ # MCP protocol layer (Unit + Integration)
├── ExcelMcp.CLI.Tests/        # CLI wrapper (Unit + Integration)
└── ExcelMcp.ComInterop.Tests/ # COM utilities (Unit + OnDemand)
```

### Test Categories

| Category | Speed | Requirements | Purpose | Run By Default |
|----------|-------|--------------|---------|----------------|
| **Unit** | Fast (2-5 sec) | None | Logic validation, no external dependencies | ✅ Yes (CI/CD safe) |
| **Integration** | Medium (10-20 min) | Excel + Windows | Real Excel COM operations, single feature focus | ✅ Yes (local development) |
| **OnDemand** | Slow (3-5 min) | Excel + Windows | Session cleanup, Excel process lifecycle | ❌ No (explicit only) |

---

## Running Tests

### By Category

```bash
# Unit tests only (CI/CD safe, no Excel)
dotnet test --filter "Category=Unit"

# Integration tests only (requires Excel)
dotnet test --filter "Category=Integration"

# OnDemand tests (session cleanup - run when modifying session/batch code)
dotnet test --filter "RunType=OnDemand"

# All tests except OnDemand (standard pre-commit)
dotnet test --filter "RunType!=OnDemand"
```

### By Speed

```bash
# Fast tests only
dotnet test --filter "Speed=Fast"

# Fast and medium speed tests (exclude slow)
dotnet test --filter "Speed=Fast|Speed=Medium"
```

### By Layer

```bash
# Core business logic tests
dotnet test --filter "Layer=Core"

# MCP Server protocol tests
dotnet test --filter "Layer=McpServer"

# CLI wrapper tests
dotnet test --filter "Layer=CLI"

# COM interop tests
dotnet test --filter "Layer=ComInterop"
```

### By Feature

```bash
# Power Query tests
dotnet test --filter "Feature=PowerQuery"

# Data Model (DAX) tests
dotnet test --filter "Feature=DataModel"

# Table tests
dotnet test --filter "Feature=Tables"

# PivotTable tests
dotnet test --filter "Feature=PivotTables"

# Range tests
dotnet test --filter "Feature=Ranges"

# VBA tests
dotnet test --filter "Feature=VBA"

# Worksheet tests
dotnet test --filter "Feature=Worksheets"

# Connection tests
dotnet test --filter "Feature=Connections"

# Parameter tests
dotnet test --filter "Feature=Parameters"
```

### Specific Test Classes

```bash
# Run specific test class
dotnet test --filter "FullyQualifiedName~PowerQueryCommandsTests"

# Run specific test method
dotnet test --filter "FullyQualifiedName~PowerQueryCommandsTests.Import_ValidMCode_CreatesQuery"
```

---

## Test Requirements & Traits

### Required Traits for All Tests

Every test MUST have these traits:

```csharp
[Trait("Category", "Unit|Integration")]      // Test category
[Trait("Speed", "Fast|Medium|Slow")]          // Execution speed
[Trait("Layer", "Core|CLI|McpServer|ComInterop")]  // Project layer
[Trait("Feature", "FeatureName")]             // Feature being tested

// Additional traits for specific test types
[Trait("RequiresExcel", "true")]              // For Integration tests
[Trait("RunType", "OnDemand")]                // For session/lifecycle tests
```

### Unit Tests

```csharp
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
public class PowerQueryValidationTests
{
    [Fact]
    public void ValidateQueryName_InvalidCharacters_ReturnsFalse()
    {
        // No Excel required - pure logic testing
    }
}
```

### Integration Tests

```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public class PowerQueryCommandsTests : IClassFixture<TempDirectoryFixture>
{
    [Fact]
    public async Task Import_ValidMCode_CreatesQuery()
    {
        // Requires Excel - real COM operations
    }
}
```

### OnDemand Tests

```csharp
[Trait("RunType", "OnDemand")]
[Trait("Speed", "Slow")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
public class ExcelSessionLifecycleTests : IAsyncLifetime
{
    [Fact]
    public async Task CreateAndDispose_NoProcessLeak()
    {
        // Tests Excel process cleanup - run explicitly only
    }
}
```

---

## Test Class Compliance Checklist

### ✅ File Organization
- [ ] Test class uses `partial` keyword for multi-file organization (if needed)
- [ ] File name matches class name (e.g., `RangeCommandsTests.cs` contains `RangeCommandsTests`)
- [ ] Files organized by feature in subdirectories (e.g., `Commands/Range/`)
- [ ] Related partial files use descriptive suffixes (e.g., `RangeCommandsTests.Values.cs`)

### ✅ Test Fixture Setup
- [ ] Test class implements `IClassFixture<TempDirectoryFixture>` for integration tests
- [ ] Constructor accepts `TempDirectoryFixture` via dependency injection
- [ ] Store temp directory in `private readonly string _tempDir` field
- [ ] **NEVER** manually implement `IDisposable` for temp directory cleanup
- [ ] **NEVER** create temp directory in constructor (fixture provides it)

### ✅ Test File Isolation
- [ ] Each test creates its own unique file using `CoreTestHelper.CreateUniqueTestFileAsync()`
- [ ] **NEVER** share a single test file across multiple tests
- [ ] **NEVER** reuse file paths between tests
- [ ] Pass `_tempDir` (from fixture) to `CreateUniqueTestFileAsync()`
- [ ] Use test class name and test method name in file creation

### ✅ File Extension Requirements
- [ ] VBA tests MUST use `.xlsm` extension (macro-enabled workbooks)
- [ ] Standard tests use `.xlsx` extension (unless VBA required)
- [ ] Pass extension parameter to `CoreTestHelper.CreateUniqueTestFileAsync()`
- [ ] **NEVER** rename files to change format (e.g., `.xlsx` → `.xlsm` fails)

### ✅ Test Assertions
- [ ] Use binary assertions: `Assert.True(result.Success, $"Reason: {result.ErrorMessage}")`
- [ ] **NEVER** use "accept both" patterns (if-success-pass, if-error-pass)
- [ ] Include descriptive failure messages in assertions
- [ ] Use `Skip` attribute if test requires unavailable features

### ✅ Integration Test Validation (Result Verification)
- [ ] **ALWAYS verify actual Excel state** after create/update operations
- [ ] **NEVER test only success status** - verify the action actually worked
- [ ] For CREATE operations: Verify object exists (list → verify it's there)
- [ ] For UPDATE operations: Verify changes persisted (view → verify formula)
- [ ] For DELETE operations: Verify object removed (list → verify gone)
- [ ] Use round-trip validation: Create/Update → Read back → Assert actual state

### ✅ Batch API Pattern
- [ ] All Core commands accept `IExcelBatch batch` as first parameter
- [ ] Tests create batch with `await ExcelSession.BeginBatchAsync(testFile)`
- [ ] Use `await using var batch` for automatic disposal
- [ ] **CRITICAL:** `await batch.SaveAsync()` MUST be called ONLY at the END of the test
- [ ] **NEVER** call `SaveAsync()` in the middle of a test
- [ ] **NEVER** call `SaveAsync()` multiple times in a single test
- [ ] Only call `SaveAsync()` if test modifies data that needs persistence verification

---

## Test Class Template

### Integration Test Template

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

        // Assert - Verify success
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        Assert.NotNull(result.Data);
        
        // Assert - Verify actual Excel state (round-trip validation)
        var verifyResult = await _commands.ListAsync(batch);
        Assert.Contains(verifyResult.Items, item => item.Name == "ExpectedName");
        
        // Save only at the end if modifications made
        await batch.SaveAsync();
    }
}
```

### Unit Test Template

```csharp
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "FeatureName")]
public class FeatureValidationTests
{
    [Theory]
    [InlineData("valid-input", true)]
    [InlineData("invalid@input", false)]
    public void ValidateInput_VariousInputs_ReturnsExpected(string input, bool expected)
    {
        // Arrange
        var validator = new FeatureValidator();

        // Act
        var result = validator.Validate(input);

        // Assert
        Assert.Equal(expected, result);
    }
}
```

---

## Common Mistakes & How to Avoid Them

### ❌ Shared Test Files
```csharp
// WRONG: Single file for all tests causes pollution
private readonly string _testFile = "SharedFile.xlsx";

// CORRECT: Unique file per test
var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
    nameof(MyTests), nameof(TestMethod), _tempDir, ".xlsx");
```

### ❌ Wrong File Extension for VBA
```csharp
// WRONG: .xlsx cannot contain VBA
var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(..., ".xlsx");

// CORRECT: .xlsm for VBA tests
var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(..., ".xlsm");
```

### ❌ SaveAsync in Middle of Test
```csharp
// WRONG: Breaks subsequent operations
var result1 = await _commands.CreateAsync(batch, "Sheet1");
await batch.SaveAsync();  // ❌ Too early!
var result2 = await _commands.RenameAsync(batch, "Sheet1", "NewName");  // FAILS!

// CORRECT: Save only at the end
var result1 = await _commands.CreateAsync(batch, "Sheet1");
var result2 = await _commands.RenameAsync(batch, "Sheet1", "NewName");
await batch.SaveAsync();  // ✅ After all operations
```

### ❌ Testing Only Success Status
```csharp
// WRONG: Doesn't verify Excel state
var result = await _commands.CreateAsync(batch, "TestTable");
Assert.True(result.Success);  // ❌ Doesn't prove table exists!

// CORRECT: Verify actual Excel state
var result = await _commands.CreateAsync(batch, "TestTable");
Assert.True(result.Success);
var listResult = await _commands.ListAsync(batch);
Assert.Contains(listResult.Tables, t => t.Name == "TestTable");  // ✅ Proves it exists!
```

### ❌ Accept Both Pattern
```csharp
// WRONG: Test always passes - feature can be 100% broken!
if (result.Success) {
    Assert.True(result.Success);
} else {
    Assert.True(result.ErrorMessage.Contains("acceptable"));
}

// CORRECT: Binary assertion
Assert.True(result.Success, $"Must succeed: {result.ErrorMessage}");
```

---

## CI/CD Configuration

### GitHub Actions (No Excel)

```yaml
- name: Run Unit Tests
  run: dotnet test --filter "Category=Unit" --logger "trx"
```

### Local Development (With Excel)

```yaml
- name: Run All Tests
  run: dotnet test --filter "RunType!=OnDemand" --logger "trx"
```

### Manual Session Tests (Requires Excel)

```bash
# Run when modifying ExcelSession.cs, ExcelBatch.cs, or ExcelHelper.cs
dotnet test --filter "RunType=OnDemand"
```

---

## Performance Targets

| Category | Test Count | Execution Time | Run Frequency |
|----------|-----------|----------------|---------------|
| **Unit** | ~46 tests | 2-5 seconds | Every build |
| **Integration** | ~150+ tests | 10-20 minutes | Pre-commit |
| **OnDemand** | ~5 tests | 3-5 minutes | Explicit only |

---

## Development Workflow

### Daily Development
```bash
# Fast feedback (2-5 seconds - excludes VBA)
dotnet test --filter "Category=Unit&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"
```

### Before Commit
```bash
# Comprehensive validation (10-20 minutes - excludes VBA)
dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"
```

### Session/Batch Code Changes
```bash
# MANDATORY when modifying session/batch code (3-5 minutes)
dotnet test --filter "RunType=OnDemand"
```

### VBA Development (Manual Only)
```bash
# Test VBA features only (requires VBA trust enabled)
dotnet test --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"
```

### Specific Feature Development
```bash
# Test only what you're working on (e.g., PowerQuery)
dotnet test --filter "Feature=PowerQuery&Category=Integration"
```

---

## Test Coverage Goals

- ✅ **Unit Tests**: 100% of business logic and validation
- ✅ **Integration Tests**: All Excel COM operations
- ✅ **Round-Trip Validation**: Create → Verify → Update → Verify → Delete → Verify
- ✅ **Error Cases**: Invalid inputs, Excel errors, edge cases

---

## Known Issues & Resolutions

### ✅ RESOLVED: Data Model Measure Creation on Reopened Files

**Issue**: `ModelMeasures.Add()` failed with "Value does not fall within expected range" on reopened files.

**Root Cause**: Microsoft documentation incorrectly states `FormatInformation` is optional. Excel requires format object for reopened files.

**Solution**: Always provide format object - use `model.ModelFormatGeneral` as default when user doesn't specify format type.

**Status**: ✅ All measure-related tests passing. Fix validated with both fresh and reopened Data Model files.

---

## Getting Help

- **Test failures**: Check test output for detailed error messages
- **Excel-specific issues**: Ensure Excel 2016+ is installed and working
- **Session/batch issues**: Run OnDemand tests to verify cleanup
- **CI/CD failures**: Ensure only Unit tests run (no Excel dependency)

---

## Summary

**Three-tier testing ensures:**
1. ✅ **Fast development** - Unit tests provide immediate feedback
2. ✅ **Excel integration** - Integration tests validate COM operations
3. ✅ **Session cleanup** - OnDemand tests prevent resource leaks

**Key principles:**
- Unique file per test (isolation)
- Round-trip validation (verify actual Excel state)
- Binary assertions (pass OR fail, never both)
- SaveAsync only at the end (prevent operation failures)

**Development workflow:**
```bash
Code → Unit tests (2-5s) → Integration tests (10-20m) → Commit
```


### GitHub Actions / Azure DevOps (No Excel Available)

```yaml
# CI environments typically don't have Excel installed
- name: Run Unit Tests
  run: dotnet test --filter "Category=Unit"
```

### Self-Hosted Runners with Excel (Optional)

```yaml
# Only if you have Windows runners with Excel installed
- name: Run Integration Tests
  run: dotnet test --filter "Category=Integration"
  # This requires Windows runners with Excel installation

- name: Run Round Trip Tests
  run: dotnet test --filter "Category=RoundTrip"
  # This requires Windows runners with Excel installation
```

### Local Development

```bash
# Quick feedback loop during development (unit tests only)
dotnet test --filter "Category=Unit"

# Feature testing with Excel (integration tests)
dotnet test --filter "Category=Integration"

# Full validation including slow round trip tests
dotnet test --filter "Category=RoundTrip"

# All non-slow tests (unit + integration)
dotnet test --filter "Speed!=Slow"
```

## Test Structure

```text
tests/
├── ExcelMcp.CLI.Tests/
│   ├── UnitTests.cs                     # [Unit, Fast] - No Excel required
│   └── Commands/
│       ├── FileCommandsTests.cs        # [Integration, Medium, Files] - Excel file operations
│       ├── PowerQueryCommandsTests.cs  # [Integration, Medium, PowerQuery] - M code automation
│       ├── ScriptCommandsTests.cs      # [Integration, Medium, VBA] - VBA script operations
│       ├── SheetCommandsTests.cs       # [Integration, Medium, Worksheets] - Sheet operations
│       └── IntegrationRoundTripTests.cs # [RoundTrip, Slow, EndToEnd] - Complex workflows
├── ExcelMcp.McpServer.Tests/
│   ├── Tools/
│   │   └── ExcelMcpServerTests.cs       # [Integration, Medium, MCP] - Direct tool method tests
│   └── Integration/
│       └── McpClientIntegrationTests.cs # [Integration, Medium, MCPProtocol] - True MCP client tests
```

## Test Organization in Test Explorer

Tests are organized using multiple traits for better filtering:

- **Category**: `Unit`, `Integration`, `RoundTrip`
- **Speed**: `Fast`, `Medium`, `Slow`
- **Feature**: `PowerQuery`, `VBA`, `Worksheets`, `Files`, `EndToEnd`

## Environment Requirements

### Unit Tests (`Category=Unit`)

- **Requirements**: None
- **Platforms**: Windows, Linux, macOS
- **CI Compatible**: ✅ Yes
- **Purpose**: Validate argument parsing, logic, validation

### Integration Tests (`Category=Integration`)

- **Requirements**: Windows + Excel installation
- **Platforms**: Windows only
- **CI Compatible**: ❌ No (unless using Windows runners with Excel)
- **Purpose**: Validate Excel COM operations, feature functionality

### MCP Protocol Tests (`Feature=MCPProtocol`)

- **Requirements**: Windows + Excel installation + Built MCP server executable
- **Platforms**: Windows only
- **CI Compatible**: ❌ No (unless using Windows runners with Excel)
- **Purpose**: True MCP client integration - starts server process and communicates via stdio

### Round Trip Tests (`Category=RoundTrip`)

- **Requirements**: Windows + Excel installation + VBA trust settings
- **Platforms**: Windows only
- **CI Compatible**: ❌ No (unless using specialized Windows runners)
- **Purpose**: End-to-end workflow validation

## MCP Testing: Tool Tests vs Protocol Tests

The MCP Server has two types of tests that serve different purposes:

### Tool Tests (`ExcelMcpServerTests.cs`)

```csharp
// Direct method calls - tests tool logic only
var result = ExcelTools.ExcelFile("create-empty", filePath);
var json = JsonDocument.Parse(result);
Assert.True(json.RootElement.GetProperty("success").GetBoolean());
```

**What it tests:**

- ✅ Tool method logic and JSON response format
- ✅ Excel COM operations and error handling
- ✅ Parameter validation and edge cases

**What it DOESN'T test:**

- ❌ MCP protocol communication (JSON-RPC over stdio)
- ❌ Tool discovery and metadata
- ❌ MCP client/server handshake
- ❌ Process lifecycle and stdio communication

### Protocol Tests (`McpClientIntegrationTests.cs`)

```csharp
// True MCP client - starts server process and communicates via stdio
var server = StartMcpServer(); // Starts actual MCP server process
var response = await SendMcpRequestAsync(server, initRequest); // JSON-RPC over stdio
```

**What it tests:**

- ✅ Complete MCP protocol implementation
- ✅ Tool discovery via `tools/list`
- ✅ JSON-RPC communication over stdio
- ✅ Server initialization and handshake
- ✅ Process lifecycle management
- ✅ End-to-end MCP client experience

**Why both are needed:**

- **Tool Tests**: Fast feedback for core functionality
- **Protocol Tests**: Validate what AI assistants actually experience

## Troubleshooting

### "Round trip tests skipped" Message

This is expected behavior. Round trip tests only run when explicitly requested:

- Use `dotnet test --filter "Category=RoundTrip"` to run them specifically
- Round trip tests are slow and not needed for regular development

### Excel COM Errors in Integration/Round Trip Tests

- **CI Environment**: Integration and Round Trip tests will fail without Excel
  - Use `--filter "Category=Unit"` in CI pipelines
  - Only run Excel-dependent tests on local machines or Windows runners with Excel
- Ensure Excel is installed (Windows only)
- Close all Excel instances before running tests
- Run `ExcelMcp setup-vba-trust` for VBA tests
- Excel COM is not available on Linux/macOS

### Slow Test Performance

- **Unit tests**: Very fast (< 1 second)
- **Integration tests**: Medium speed (5-15 seconds)
- **Round trip tests**: Very slow (30+ seconds)
- **CI Strategy**: Run only unit tests in CI (no Excel required)
- **Local Development**:
  - Use unit tests for rapid development cycles
  - Use integration tests for feature validation
  - Use round trip tests for comprehensive end-to-end validation

## Adding New Tests

### Fast Unit Test

```csharp
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
public class MyUnitTests
{
    [Fact]
    public void MyMethod_WithValidInput_ReturnsExpected()
    {
        // Test logic without Excel COM
    }
}
```

### Integration Test

```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "YourFeature")]  // e.g., "PowerQuery", "VBA", "Worksheets", "Files"
public class MyCommandsTests : IDisposable
{
    [Fact]
    public void MyCommand_WithExcel_WorksCorrectly()
    {
        // Test logic using Excel COM automation
    }
}
```

### Round Trip Test

```csharp
[Trait("Category", "RoundTrip")]
[Trait("Speed", "Slow")]
[Trait("Feature", "EndToEnd")]
public class MyRoundTripTests : IDisposable
{
    [Fact]
    public void ComplexWorkflow_EndToEnd_WorksCorrectly()
    {
        // Complex workflow testing multiple features together
    }
}
```

## Benefits of This Test Organization

✅ **Fast feedback during development** (unit tests)  
✅ **Feature validation with Excel** (integration tests)  
✅ **Comprehensive end-to-end validation when requested** (round trip tests)  
✅ **Flexible filtering by category, speed, or feature**  
✅ **Better organization in Test Explorer**  
✅ **CI/CD flexibility** (different test suites for different scenarios)  
✅ **Clear documentation for contributors**
