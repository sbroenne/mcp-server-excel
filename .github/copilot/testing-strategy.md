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

## Key Lessons

1. **OnDemand pattern** - Essential for Excel-dependent tests that can't run in CI/CD
2. **Test isolation** - Save/restore global state (like `ExcelHelper.InstancePool`)
3. **Realistic data** - Use test helpers to create real Excel objects
4. **Layer prefixes** - Prevent FQDN conflicts in test class names
5. **Complete traits** - All tests MUST have Category, Speed, Layer traits
