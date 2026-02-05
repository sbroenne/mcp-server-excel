# ExcelMcp Tests

> **⚠️ No Traditional Unit Tests**: ExcelMcp has no unit tests. Integration tests ARE our unit tests because Excel COM cannot be meaningfully mocked. See [`docs/ADR-001-NO-UNIT-TESTS.md`](../docs/ADR-001-NO-UNIT-TESTS.md) for full architectural rationale.

## Quick Start

```powershell
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
├── ExcelMcp.Core.Tests/           # Core business logic (Integration)
├── ExcelMcp.Diagnostics.Tests/    # Excel COM behavior research (OnDemand, Manual)
├── ExcelMcp.McpServer.Tests/      # MCP protocol layer (Integration)
├── ExcelMcp.CLI.Tests/            # CLI wrapper (Integration)
└── ExcelMcp.ComInterop.Tests/     # COM utilities (OnDemand)

llm-tests/                          # LLM tool behavior validation (Manual)
```

## Test Categories

| Category | Speed | Requirements | Run By Default |
|----------|-------|--------------|----------------|
| **Integration** | Medium (10-20 min) | Excel + Windows | ✅ Yes (local) |
| **OnDemand** | Slow (3-5 min) | Excel + Windows | ❌ No (explicit only) |
| **Diagnostics** | Slow (varies) | Excel + Windows | ❌ No (manual, excluded from CI) |
| **LLM Tests** | Slow (varies) | Excel + Azure OpenAI | ❌ No (manual only) |

## Diagnostics Tests

Diagnostics tests are research/exploratory tests in `ExcelMcp.Diagnostics.Tests` that document the actual behavior of Excel's COM APIs without our abstraction layer. These tests are **excluded from CI** to keep automation focused on core functionality.

**Purpose:**
- Understand Excel COM API behavior for Power Query, Data Model, PivotTables, etc.
- Document findings and edge cases for future implementation decisions
- Test alternative approaches to complex Excel operations

**Trait markers:**
- `Layer=Diagnostics`  
- `RunType=OnDemand`

**Run diagnostics tests locally:**
```powershell
# All diagnostics tests
dotnet test tests/ExcelMcp.Diagnostics.Tests/ --filter "RunType=OnDemand&Layer=Diagnostics"

# Specific diagnostic tests
dotnet test tests/ExcelMcp.Diagnostics.Tests/ --filter "Feature=PowerQuery&RunType=OnDemand"
```

**CI Behavior:**
- Diagnostics tests are **NOT** run in CI workflows (GitHub Actions)
- Path filter includes folder to trigger builds when tests change
- Test execution uses `RunType!=OnDemand` filter to exclude them

## Feature-Specific Tests

```powershell
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
| **Daily development** | `dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA"` |
| **Before commit** | `dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA"` |
| **Modified session/batch code** | `dotnet test --filter "RunType=OnDemand"` (see [Rule 3](../.github/instructions/critical-rules.instructions.md#rule-3-session-cleanup-tests)) |
| **VBA development** | `dotnet test --filter "(Feature=VBA\|Feature=VBATrust)&RunType!=OnDemand"` |
| **LLM behavior validation** | See [LLM Tests](#llm-tests) section below |

## LLM Tests

The `llm-tests/` project validates that LLMs correctly use Excel MCP Server and CLI tools using [pytest-aitest](https://github.com/sbroenne/pytest-aitest).

### When to Run LLM Tests

- **Manual/on-demand only** - Not part of CI/CD
- After changing tool descriptions or adding new tools
- To validate LLM behavior patterns (e.g., incremental updates vs rebuild)

### Running LLM Tests

```powershell
# From llm-tests/
uv sync
uv run pytest -m aitest -v
```

### Prerequisites

- `AZURE_OPENAI_ENDPOINT` environment variable
- Windows desktop with Excel installed
- pytest-aitest dependency (local path via uv)

**See [LLM Tests README](../llm-tests/README.md) for complete documentation.**

## VBA Testing

### Why VBA Tests Are Excluded by Default

VBA tests are excluded from normal test runs because:
1. **Stable codebase** - VBA features are mature with minimal changes
2. **Performance** - Excluding VBA tests makes integration tests ~25% faster (10-15 min vs 15-20 min)
3. **Special requirements** - VBA tests require VBA trust enabled in Excel settings
4. **Opt-in model** - Explicit testing when VBA code changes, rather than every commit

### When to Run VBA Tests

Run VBA tests manually when:
- Modifying VBA-related code (ScriptCommands, VbaTrustDetection)
- Adding new VBA features
- Before releasing VBA-related changes
- Troubleshooting VBA-specific issues

### How to Run VBA Tests

```powershell
# Run ONLY VBA tests
dotnet test --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"

# Run ALL tests including VBA (takes longer)
dotnet test --filter "Category=Integration&RunType!=OnDemand"
```

### VBA Test Files

All VBA tests are tagged with `[Trait("Feature", "VBA")]` or `[Trait("Feature", "VBATrust")]`:

```
tests/ExcelMcp.Core.Tests/Integration/Commands/Script/
  - ScriptCommandsTests.cs
  - ScriptCommandsTests.Lifecycle.cs
  - VbaTrustDetectionTests.ScriptCommands.cs
  - VbaTrustDetectionTests.cs

tests/ExcelMcp.CLI.Tests/Integration/Commands/
  - ScriptAndSetupCommandsTests.cs
```

### VBA Trust Setup

VBA tests require VBA trust enabled in Excel:

```powershell
# Enable VBA trust (required for VBA tests)
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "AccessVBOM" -Value 1

# Verify setting
Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "AccessVBOM"
```

**Security Note:** Only enable VBA trust in development environments. Production systems should keep this disabled.

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
