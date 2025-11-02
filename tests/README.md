# ExcelMcp Tests

> **⚠️ No Traditional Unit Tests**: ExcelMcp has no unit tests. Integration tests ARE our unit tests because Excel COM cannot be meaningfully mocked. See [`docs/ADR-001-NO-UNIT-TESTS.md`](../docs/ADR-001-NO-UNIT-TESTS.md) for full architectural rationale.

## Quick Start

```bash
# Development (fast feedback - excludes VBA tests)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Pre-commit (comprehensive - excludes VBA tests)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Session/batch changes (MANDATORY when modifying session/batch code)
dotnet test --filter "RunType=OnDemand"

# VBA tests (manual only - requires VBA trust enabled)
dotnet test --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"
```

## Documentation

**For complete testing guidance, see:**

- **[Testing Strategy](../.github/instructions/testing-strategy.instructions.md)** - Quick reference, templates, common mistakes
- **[Critical Rules](../.github/instructions/critical-rules.instructions.md)** - Mandatory development rules (Rule 14: SaveAsync)

## Test Architecture

```
tests/
├── ExcelMcp.Core.Tests/      # Core business logic (Unit + Integration)
├── ExcelMcp.McpServer.Tests/ # MCP protocol layer (Unit + Integration)
├── ExcelMcp.CLI.Tests/        # CLI wrapper (Unit + Integration)
└── ExcelMcp.ComInterop.Tests/ # COM utilities (Unit + OnDemand)
```

## Test Categories

| Category | Speed | Requirements | Run By Default |
|----------|-------|--------------|----------------|
| **Unit** | Fast (2-5 sec) | None | ✅ Yes (CI/CD) |
| **Integration** | Medium (10-20 min) | Excel + Windows | ✅ Yes (local) |
| **OnDemand** | Slow (3-5 min) | Excel + Windows | ❌ No (explicit only) |

## Feature-Specific Tests

```bash
# Test specific feature only
dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"
dotnet test --filter "Feature=DataModel&RunType!=OnDemand"
dotnet test --filter "Feature=Tables&RunType!=OnDemand"
dotnet test --filter "Feature=PivotTables&RunType!=OnDemand"
dotnet test --filter "Feature=Ranges&RunType!=OnDemand"
dotnet test --filter "Feature=Connections&RunType!=OnDemand"
```

## When to Run Which Tests

| Scenario | Command |
|----------|---------|
| **Daily development** | `dotnet test --filter "Category=Unit&RunType!=OnDemand&Feature!=VBA"` |
| **Before commit** | `dotnet test --filter "(Category=Unit\|Category=Integration)&RunType!=OnDemand&Feature!=VBA"` |
| **Modified session/batch code** | `dotnet test --filter "RunType=OnDemand"` (see [Rule 3](../.github/instructions/critical-rules.instructions.md#rule-3-session-cleanup-tests)) |
| **VBA development** | `dotnet test --filter "(Feature=VBA\|Feature=VBATrust)&RunType!=OnDemand"` |

## Key Principles

- ✅ **File Isolation** - Each test creates unique file (no sharing)
- ✅ **Binary Assertions** - Pass OR fail, never "accept both"
- ✅ **Verify Excel State** - Always verify actual Excel state after operations
- ❌ **No SaveAsync** - Unless testing persistence (see [Rule 14](../.github/instructions/critical-rules.instructions.md#rule-14-no-saveasync-unless-testing-persistence))

## Getting Help

- **Test failures**: Check test output for detailed error messages
- **Excel issues**: Ensure Excel 2016+ installed and activated
- **Session/batch issues**: Run OnDemand tests to verify cleanup
- **Writing tests**: See [Testing Strategy](../.github/instructions/testing-strategy.instructions.md)
