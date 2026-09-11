# ExcelMcp Tests

Excel-dependent behavior uses real Excel integration tests. Parsing, mapping,
serialization, and generation can use focused tests without Excel. The former
blanket ban on unit tests is [superseded](../docs/ADR-001-NO-UNIT-TESTS.md).

## Quick Start

```powershell
# One Core feature
dotnet test tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj --filter "Feature=PowerQuery&RunType!=OnDemand"

# Excel-independent parsing
dotnet test tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj --filter "FullyQualifiedName~ServiceRegistryJsonParsingTests"

# Session/batch changes: narrow by test name when appropriate
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"

# VBA behavior (requires VBA trust enabled)
dotnet test tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"
```

Set a hard execution timeout for every Excel-dependent run. Run only the
relevant project and filter, not the full integration suite during iteration.

## Documentation

**For complete testing guidance, see:**

- **[Testing Strategy](../.github/instructions/testing-strategy.instructions.md)** - Quick reference, templates, common mistakes
- **[Repository Rules](../.github/copilot-instructions.md)** - Build, E2E, and contribution requirements

## Test Architecture

```
tests/
├── ExcelMcp.Core.Tests/           # Excel behavior and pure parsing tests
├── ExcelMcp.Diagnostics.Tests/    # Excel COM behavior research (OnDemand, Manual)
├── ExcelMcp.McpServer.Tests/      # MCP protocol layer (Integration)
├── ExcelMcp.CLI.Tests/            # CLI wrapper (Integration)
├── ExcelMcp.ComInterop.Tests/     # COM utilities and session infrastructure
└── ExcelMcp.SkillGeneration.Tests/ # Generated skill and plugin checks

llm-tests/                          # LLM tool behavior validation (Manual)
```

## Test Categories

| Category | Speed | Requirements | Run By Default |
|----------|-------|--------------|----------------|
| **Unit** | Fast | No Excel for pure logic | Select the relevant tests |
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
| **Daily development** | Run the smallest project and feature/name filter covering the change. |
| **Before commit** | Rerun affected tests and applicable checks; follow the [runtime E2E requirements](../.github/copilot-instructions.md#build-and-validation). |
| **Modified session/batch code** | Run relevant OnDemand tests in `ExcelMcp.ComInterop.Tests`; see [Testing Strategy](../.github/instructions/testing-strategy.instructions.md#commands). |
| **VBA development** | `dotnet test --filter "(Feature=VBA\|Feature=VBATrust)&RunType!=OnDemand"` |
| **LLM behavior validation** | See [LLM Tests](#llm-tests) section below |

## LLM Tests

The `llm-tests/` project validates that LLMs correctly use Excel MCP Server and CLI tools using [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering).

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
- GitHub auth via `gh auth login` or `GITHUB_TOKEN`

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

All VBA tests are tagged with `[Trait("Feature", "VBA")]`:

```
tests/ExcelMcp.Core.Tests/Integration/Commands/Vba/
  - VbaCommandsTests.cs
  - VbaCommandsTests.Trust.cs
  - VbaCommandsTests.Trust.ScriptCommands.cs

tests/ExcelMcp.CLI.Tests/Integration/
  - VbaRunValidationCliTests.cs
  - VbaRunCliTransportProofTests.cs
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

### Designing a regression test

Reproduce the reported failure before changing the implementation. Cover
meaningful boundary and error cases, plus both entry points when the contract
crosses CLI and MCP. There is no fixed test quota.

Assert resulting workbook state and relevant returned fields rather than only
`Success`. For update/replace behavior, assert both that old content is absent
and that new content is exact. Error assertions should distinguish the intended
failure from other exceptions instead of accepting incompatible outcomes.

Use a unique workbook with the established feature fixture. Combining
`IClassFixture<T>` and a collection fixture on the same class can create competing
Excel sessions. Follow neighboring trait conventions and the COM cleanup rules
for any references acquired by the test itself.

For in-memory changes, inspect the same batch without saving. For persistence,
save and close, reopen in a new batch, then assert the state; do not open the
same workbook in two live batches. Use `.xlsm` for VBA persistence.

### Diagnosing a failing test

Run the failure alone before broadening the run. Check workbook isolation,
fixture selection, actual Excel state, cleanup, and whether the assertion needs
a save/reopen cycle. Inspect fallback/retry paths when a primary-path fix is
insufficient. Do not hide a deterministic failure with skip/xfail or loosen an
assertion merely to make it pass.

- ✅ **File Isolation** - Each test creates unique file (no sharing)
- ✅ **Binary Assertions** - Pass OR fail, never "accept both"
- ✅ **Verify Excel State** - Always verify actual Excel state after operations
- **Explicit persistence** - Call `batch.Save()` only when testing save/close/reopen behavior (see [Testing Strategy](../.github/instructions/testing-strategy.instructions.md#save-and-round-trip-behavior)).

## Getting Help

- **Test failures**: Check test output for detailed error messages
- **Excel issues**: Ensure Excel 2016+ installed and activated
- **Session/batch issues**: Run OnDemand tests to verify cleanup
- **Writing tests**: See [Testing Strategy](../.github/instructions/testing-strategy.instructions.md)
