---
name: "project-conventions"
description: "Core conventions and patterns for ExcelMcp"
domain: "project-conventions"
confidence: "high"
source: "configured"
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

## Anti-Patterns

- **Unit tests** — NEVER write unit tests. Integration tests only for COM interop.
- **RefreshAll()** — NEVER use. Use individual `queryTable.Refresh(false)` (synchronous).
- **Catch-and-swallow** — NEVER catch exceptions in Core commands to return error results.
- **Dual test fixtures** — NEVER use both `IClassFixture<T>` AND `[Collection("...")]`.
- **Manual ScreenUpdating** — ExcelWriteGuard handles this automatically.
- **Suppressing EnableEvents** — Data Model operations depend on them.
