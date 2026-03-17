---
name: "project-conventions"
description: "Core conventions and patterns for ExcelMcp"
---

## Context

ExcelMcp is a Windows-only toolset for programmatic Excel automation via COM interop, with TWO equal entry points: MCP Server (for AI assistants) and CLI (for scripting/agents). .NET/C# codebase.

## Patterns

### Exception Propagation (CRITICAL)

Core commands NEVER wrap `batch.Execute()` in try-catch that returns error results. Let exceptions propagate naturally — `batch.Execute()` handles them via `TaskCompletionSource`.

```csharp
// ✅ CORRECT
return await batch.Execute((ctx, ct) => {
    // operation
    return ValueTask.FromResult(new OperationResult { Success = true });
});

// ❌ WRONG — double-wraps, loses stack context
try { return await batch.Execute(...); }
catch (Exception ex) { return new OperationResult { Success = false, ErrorMessage = ex.Message }; }
```

### COM Object Cleanup

ALL dynamic COM objects must be released in `finally` blocks using `ComUtilities.Release(ref obj!)`. NEVER use catch blocks to swallow exceptions.

### Success Flag Invariant

`Success == true` ⟹ `ErrorMessage == null || ErrorMessage == ""`. Set Success in try block, always false in catch.

### Error Handling

- Core Commands: Let exceptions propagate through batch.Execute()
- MCP Server: Return JSON with `isError: true` for business errors; throw McpException for validation
- CLI: Wrap Core calls in try-catch, display with `AnsiConsole.MarkupLine`

### Testing

- Framework: xUnit with integration tests ONLY (no unit tests)
- Test location: `tests/ExcelMcp.Core.Tests/`, `tests/ExcelMcp.McpServer.Tests/`, `tests/ExcelMcp.CLI.Tests/`
- Run command: `dotnet test tests/ExcelMcp.Core.Tests --filter "Feature=<name>&RunType!=OnDemand"`
- TDD: Write test FIRST → RED → implement → GREEN
- NEVER share test files between tests — each test creates unique files
- ALWAYS verify actual Excel state, not just success flags (round-trip validation)
- For range write bugs, test payload shape explicitly: rectangular wide writes, jagged rows, and create-sheet-then-write-non-A1 flows. Don’t infer COM limits from jagged input failures.

### Bug Report Triage For Tests

- Check the live tool surface before treating a report as a missing-feature bug; verify `ServiceAction` coverage and current MCP tool docs first.
- Classify each report item before writing tests: regression in promised behavior, discoverability/documentation gap, or new feature request.
- Regressions get exact failing workflow tests first at Core and MCP layers.
- Existing capabilities with weak discoverability get positive end-to-end coverage before any API expansion.
- New features get acceptance tests only after the public API shape is agreed.

### Code Style

- Analyzer: `TreatWarningsAsErrors=true` with .NET analyzers
- Naming: PascalCase for public, camelCase for params → auto-converts to snake_case in MCP
- One public class per file, file name = class name
- Partial classes for 15+ methods (split by feature domain)
- No emojis in LLM-consumed content (XML docs, skill .md files)

### File Structure

```
src/ExcelMcp.ComInterop/  — COM patterns, STA threading, sessions
src/ExcelMcp.Core/         — Excel business logic, commands
src/ExcelMcp.Service/      — Session management, command routing
src/ExcelMcp.McpServer/    — MCP protocol tools
src/ExcelMcp.CLI/          — Command-line interface
src/ExcelMcp.Generators*/  — Source generators
tests/                     — Integration tests (no unit tests)
skills/shared/             — Single source of truth for docs/prompts
```

### Formatting Surface Split

Before classifying a formatting bug as missing functionality, check both `range` and `range_format`.

- `range` owns value/formula-adjacent display formats such as `set-number-format`
- `range_format` owns visual styling and layout actions such as `format-range`, `auto-fit-columns`, and `auto-fit-rows`

If the capability already exists under one of those tools, treat it as a discoverability/API-shape issue first, not a backend feature gap.

For real batching work, prefer existing list-of-objects patterns over ad hoc JSON blobs. `IRangeEditCommands.Sort(... List<SortColumn> ...)` is a precedent that the shared surface can carry structured collections cleanly.

## Anti-Patterns

- **Unit tests** — NEVER write unit tests. Integration tests only for COM interop.
- **RefreshAll()** — NEVER use. Use individual `queryTable.Refresh(false)` (synchronous).
- **Catch-and-swallow** — NEVER catch exceptions in Core commands to return error results.
- **Dual test fixtures** — NEVER use both `IClassFixture<T>` AND `[Collection("...")]`.
- **Manual ScreenUpdating** — ExcelWriteGuard handles this automatically.
- **Suppressing EnableEvents** — Data Model operations depend on them.
- **Assuming Excel has a hidden range-width cap** — if `set-values` fails with `ArgumentOutOfRangeException`, inspect for jagged `List<List<object?>>` input before blaming COM.

## Triage Pattern

Before assigning a bug to Core, check three things in order:

1. Existing integration coverage for the exact shape or a close analogue.
2. Whether the capability already exists under a different tool or action name.
3. Whether the failure is more likely in MCP/service argument binding, docs/skills discoverability, or true COM/Core behavior.

Use this especially for reports that claim a hard product limit or a missing feature. Wide-range failures and formatting gaps are often mis-triaged when tests or tool surfaces already cover the scenario elsewhere.
