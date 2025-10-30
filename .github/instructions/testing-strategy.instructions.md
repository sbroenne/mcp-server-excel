---
applyTo: "tests/**/*.cs"
---

# Testing Strategy

> **Three-tier testing: Unit (fast) → Integration (Excel) → OnDemand (pool cleanup)**

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

// OnDemand Tests (pool cleanup, stress tests)
[Trait("RunType", "OnDemand")]
[Trait("Speed", "Slow")]
```

## Development Workflow

```bash
# Fast feedback during development
dotnet test --filter "Category=Unit&RunType!=OnDemand"

# Pre-commit validation
dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand"

# Pool code changes (MANDATORY - see CRITICAL-RULES.md)
dotnet test --filter "RunType=OnDemand"
```

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
- ✅ Modifying `ExcelInstancePool.cs` or `ExcelHelper.cs`
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

