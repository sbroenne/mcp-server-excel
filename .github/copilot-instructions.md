# GitHub Copilot Instructions - ExcelMcp

> **🎯 Optimized for AI Coding Agents** - Modular, path-specific instructions

## 📋 Critical Files (Read These First)

**ALWAYS read when working on code:**
- [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) - 27 mandatory rules (Success flag, COM cleanup, tests, etc.)
- [Architecture Patterns](instructions/architecture-patterns.instructions.md) - Batch API, command pattern, resource management

**Read based on task type:**
- Adding/fixing commands → [Excel COM Interop](instructions/excel-com-interop.instructions.md)
- Writing tests → [Testing Strategy](instructions/testing-strategy.instructions.md)
- MCP Server work → [MCP Server Guide](instructions/mcp-server-guide.instructions.md)
- Creating PR → [Development Workflow](instructions/development-workflow.instructions.md)
- Fixing bugs → [Bug Fixing Checklist](instructions/bug-fixing-checklist.instructions.md)

**Less frequently needed:**
- [Excel Connection Types](instructions/excel-connection-types-guide.instructions.md) - Only for connection-specific work
- [README Management](instructions/readme-management.instructions.md) - Only when updating READMEs
- [Documentation Structure](instructions/documentation-structure.instructions.md) - Only when creating docs

---

## What is ExcelMcp?

**ExcelMcp** is a Windows-only toolset for programmatic Excel automation via COM interop, designed for coding agents and automation scripts.

> **⚠️ CRITICAL: ExcelMcp has TWO equal entry points — MCP Server AND CLI.**
> Both are first-class citizens. Every feature, action, and parameter must work identically through both.
> When adding/changing features, ALWAYS verify BOTH MCP Server tools AND CLI commands are updated.
> See Rule 24 (Post-Change Sync) for the full checklist.

**Core Layers:**
1. **ComInterop** (`src/ExcelMcp.ComInterop`) - Reusable COM automation patterns (STA threading, session management, batch operations, OLE message filter)
2. **Core** (`src/ExcelMcp.Core`) - Excel-specific business logic (Power Query, VBA, worksheets, parameters)
3. **Service** (`src/ExcelMcp.Service`) - Excel session management and command routing (in-process for MCP Server, named pipe for CLI daemon)
4. **CLI** (`src/ExcelMcp.CLI`) - Command-line interface for scripting (EQUAL entry point)
5. **MCP Server** (`src/ExcelMcp.McpServer`) - Model Context Protocol for AI assistants (EQUAL entry point)

**Source Generators** (`src/ExcelMcp.Generators*`) - Generate CLI commands and MCP tools from Core interfaces

---

## 🎯 Quick Reference

### Test Commands
```powershell
# ⚠️ CRITICAL: Integration tests take 45+ MINUTES for full suite
# ALWAYS use surgical testing - test only what you changed!
# ALWAYS run tests with an explicit timeout in the terminal/tooling layer.
# Never leave test runs open-ended; fail fast if Excel or COM automation stalls.

# Fast feedback (excludes VBA) - Still takes 10-15 minutes
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Surgical testing - Feature-specific (2-5 minutes per feature)
dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"
dotnet test --filter "Feature=Ranges&RunType!=OnDemand"
dotnet test --filter "Feature=PivotTables&RunType!=OnDemand"

# Session/batch changes (MANDATORY)
dotnet test --filter "RunType=OnDemand"
```

### Code Patterns
```csharp
// Core: NEVER wrap batch.Execute() in try-catch that returns error result
// Let exceptions propagate naturally - batch.Execute() handles them via TaskCompletionSource
public DataType Method(IExcelBatch batch, string arg1)
{
    return batch.Execute((ctx, ct) => {
        dynamic? item = null;
        try {
            // Operation code here
            item = ctx.Book.SomeObject;
            // For CRUD: return void (throws on error)
            // For queries: return actual data
            return someData;
        }
        finally {
            // ✅ ONLY finally blocks for COM cleanup
            ComUtilities.Release(ref item!);
        }
        // ❌ NO catch blocks that return error results
    });
}


// CLI: Wrap Core calls
public int Method(string[] args)
{
    try {
        using var batch = ExcelSession.BeginBatch(filePath);
        _coreCommands.Method(batch, arg1);
        return 0;
    } catch (Exception ex) {
        AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
        return 1;
    }
}

// Tests: Use batch API
[Fact]
public void TestMethod()
{
    using var batch = ExcelSession.BeginBatch(_testFile);
    var result = _commands.Method(batch, args);
    Assert.NotNull(result); // Or other appropriate assertion
}
```

### Tool Selection
- Code changes → `replace_string_in_file` (3-5 lines context)
- Find code → `grep_search` or `semantic_search`
- Check errors → `get_errors`
- Build/test/git → `run_in_terminal`

---

## 🔄 Key Lessons (Update After Major Work)

**Success Flag:** NEVER `Success = true` with `ErrorMessage`. Set Success in try block, always false in catch.

**Batch API:** Create NEW simple tests. CLI needs try-catch wrapping.

**Excel Quirks:** Type 3/4 both handle TEXT. `RefreshAll()` unreliable. Use `queryTable.Refresh(false)`.

**MCP Design:** Prompts are shortcuts, not tutorials. LLMs know Excel/programming.

**Tool Priority:** `replace_string_in_file` > `grep_search` > `run_in_terminal`. Avoid PowerShell for code.

**Pre-Commit:** Search TODO/FIXME/HACK, delete commented code, verify tests, check docs.

**PR Review:** Check automated comments immediately (Copilot, GitHub Security). Fix before human review.

**Surgical Testing:** Integration tests take 45+ minutes. ALWAYS test only the feature you changed using `--filter "Feature=<name>"`.

**Test Timeouts:** ALWAYS set an explicit timeout when running tests from terminal or agent tooling so hung Excel/COM runs fail fast instead of blocking the session.

**MCP Parameter Naming:** NEVER use underscores in C# Core interface parameter names. The `McpToolGenerator` calls `StringHelper.ToSnakeCase()` on the C# parameter name to produce the MCP snake_case parameter automatically. Use camelCase in C# that produces the desired snake_case output: `rangeAddress` → `range_address`, `sourceRangeAddress` → `source_range_address`. If the C# name can't produce the desired MCP name via ToSnakeCase, use `[FromString("desiredName")]` attribute instead of underscores in C# names.

**ExcelWriteGuard (Structural Safety):** `Execute()` automatically suppresses `ScreenUpdating` via `ExcelWriteGuard`. Do NOT add `ScreenUpdating` suppression in command code. Calculation suppression is manual and ONLY in value/formula write commands (SetValues, SetFormulas, Append, Write). NEVER suppress `EnableEvents` or `Calculation` universally — Data Model, PivotTable, and Power Query operations depend on them.

**Shutdown Resilience:** ALL workbook close paths (single AND multi-workbook) use `ExcelShutdownService`. Save and Close have retry for transient errors. PID capture has retry for Hwnd=0. `AppDomain.ProcessExit` handler kills tracked Excel PIDs on crash.

**Test Fixture Anti-Pattern:** NEVER use both `IClassFixture<T>` and `[Collection("...")]` with a collection fixture on the same test class. Dual fixtures create concurrent Excel sessions that deadlock during initialization with `maxParallelThreads: 1`. Use ONLY the collection fixture.

**Golden Rule (Diagnose Before Coding):** No changes without a failing test first. Write a test that proves the bug exists, watch it fail, then fix it, then watch it pass. Diagnose root cause before writing any code — spent a full session implementing the wrong fix once (issue was daemon dying, but coded a UI fix + 4 tests, all reverted).

**COM Fix Patterns (2026):**
- `OleMessageFilter.MessagePending` must return `WAITDEFPROCESS` (1), not `WAITNOPROCESS` (2) — causes STA deadlock on re-entrant COM callbacks (e.g. conditional formatting on formula cells).
- `OleMessageFilter.RetryRejectedCall` must retry `SERVERCALL_REJECTED` (dwRejectType=1) for 120s — enterprise auth dialogs cause repeated rejections.
- `ExcelBatch` starts Excel visible during open so auth/sign-in dialogs are interactable, hides after success. Tests suppress via `ExcelBatch.SuppressVisibleDuringOpen = true` in `[ModuleInitializer]`.
- VBA `App.Run()` must use late-bound `Type.InvokeMember("Run", BindingFlags.InvokeMethod, ...)` — early-bound PIA `Run()` triggers `FileNotFoundException` for `Microsoft.Vbe.Interop.dll` when VBE isn't installed.
- Startup leak fix: use locals (`startupExcel`, `startupPrimaryWorkbook`, `startupWorkbooks`) for COM cleanup during `ExcelBatch` open — don't rely on session fields that may not be set if open fails mid-way.
- `ComDiagnostics.Collect()` utility gathers COM environment info (CLSID, PIA GUID, bitness, Office install type) for enriching `InvalidCastException` messages.
- `WithSessionAsync` must catch `OperationCanceledException` and force-close session; `ExcelBatch.Execute` must fail fast if `_operationTimedOut` is set.

**GitHub Issue Comments:** ALWAYS verify @mention usernames match the actual issue/PR author before posting. Read the issue/PR to confirm the author's handle — wrong @mentions are embarrassing and unprofessional.

---

## 📚 How Path-Specific Instructions Work

GitHub Copilot auto-loads instructions based on files you're editing:

- `tests/**/*.cs` → [Testing Strategy](instructions/testing-strategy.instructions.md)
- `src/ExcelMcp.Core/**/*.cs` → [Excel COM Interop](instructions/excel-com-interop.instructions.md)
- `src/ExcelMcp.McpServer/**/*.cs` → [MCP Server Guide](instructions/mcp-server-guide.instructions.md)
- `.github/workflows/**/*.yml` → [Development Workflow](instructions/development-workflow.instructions.md)
- `**` (all files) → [CRITICAL-RULES.md](instructions/critical-rules.instructions.md)

Modular approach = relevant context without overload.

---

## 🔒 Pre-Commit Hooks (10 Automated Checks)

Pre-commit runs `scripts/pre-commit.ps1` which blocks commits if any check fails:

| # | Check | Script | What It Validates |
|---|-------|--------|-------------------|
| 1 | Branch | (inline) | Never commit to `main` directly (Rule 6) |
| 2 | COM Leaks | `check-com-leaks.ps1` | All `dynamic` COM objects have `ComUtilities.Release()` in finally |
| 3 | Coverage Audit | `audit-core-coverage.ps1` | 100% Core methods exposed via MCP Server |
| 4 | MCP-Core Implementation | `check-mcp-core-implementations.ps1` | All enum actions have Core method implementations |
| 5 | Success Flag | `check-success-flag.ps1` | Rule 0: Never `Success=true` with `ErrorMessage` |
| 6 | CLI Coverage | `check-cli-coverage.ps1` | All action enums have CLI commands |
| 7 | CLI Action Switch | `check-cli-action-coverage.ps1` | Actions requiring args have explicit switch cases |
| 8 | CLI Settings Usage | `check-cli-settings-usage.ps1` | All Settings properties used in args |
| 9 | CLI Workflow Test | `Test-CliWorkflow.ps1` | E2E CLI workflow smoke test |
| 10 | MCP Smoke Test | `dotnet test --filter "...SmokeTest..."` | All MCP tools functional |

**Install hook:**
```powershell
# From repo root
Copy-Item scripts\pre-commit.ps1 .git\hooks\pre-commit
```

---

## 🧪 LLM Integration Tests (`llm-tests/`)

Separate pytest-based project validating LLM behavior using `pytest-aitest`:

```powershell
# Setup
cd llm-tests
uv sync

# Run tests
uv run pytest -m mcp -v      # MCP Server tests
uv run pytest -m cli -v      # CLI tests
uv run pytest -m aitest -v   # All LLM tests
```

**Prerequisites:**
- Azure OpenAI endpoint: `$env:AZURE_OPENAI_ENDPOINT = "https://<resource>.openai.azure.com/"`
- Build MCP Server: `dotnet build src\ExcelMcp.McpServer -c Release`

**Structure:**
- `test_mcp_*.py` - MCP Server workflows
- `test_cli_*.py` - CLI workflows
- `Fixtures/` - Shared test inputs (CSV/JSON/M files)

---

## 📦 Agent Skills (`skills/`)

Two cross-platform AI assistant skill packages:

| Skill | File | Target | Best For |
|-------|------|--------|----------|
| **excel-cli** | `skills/excel-cli/SKILL.md` | CLI Tool | Coding agents (token-efficient, `--help` discoverable) |
| **excel-mcp** | `skills/excel-mcp/SKILL.md` | MCP Server | Conversational AI (rich tool schemas) |

**Build skills from source:**
```powershell
dotnet build -c Release  # Generates SKILL.md, copies references, and generates MCP prompts
```

**Guidance architecture (single source of truth):**
- `skills/shared/*.md` → auto-copied to skill references AND auto-generated as MCP prompts
- Skill-based clients (VS Code, Cursor) read `skills/excel-*/references/`
- MCP-only clients (Claude Desktop) read auto-generated `[McpServerPrompt]` methods
- NEVER create separate prompt files for content that belongs in `skills/shared/`

**Install via npx:**
```bash
npx skills add sbroenne/mcp-server-excel --skill excel-cli   # Coding agents
npx skills add sbroenne/mcp-server-excel --skill excel-mcp   # Conversational AI
```

---

## 🏗️ Architecture Patterns

### Command File Structure
```
Commands/Sheet/
├── ISheetCommands.cs           # Interface (defines contract)
├── SheetCommands.cs            # Partial class (constructor, DI)
├── SheetCommands.Lifecycle.cs  # Partial (Create, Delete, Rename...)
└── SheetCommands.Style.cs      # Partial (formatting operations)
```

**Rules:**
- One public class per file
- File name = class name
- Partial classes for 15+ methods (split by feature domain)

### Exception Propagation (CRITICAL)
```csharp
// ✅ CORRECT: Let batch.Execute() handle exceptions
return await batch.Execute((ctx, ct) => {
    var result = DoSomething();
    return ValueTask.FromResult(result);
});
// Exception auto-caught by TaskCompletionSource → OperationResult { Success = false }

// ❌ WRONG: Never suppress with catch returning error result
catch (Exception ex) { 
    return new OperationResult { Success = false, ErrorMessage = ex.Message }; 
}
```

### Service Architecture (TWO EQUAL ENTRY POINTS)

```
MCP Server ──► In-process ExcelMcpService ──► Core Commands ──► Excel COM
CLI ─────────► CLI Daemon (named pipe) ─────► Core Commands ──► Excel COM
```

**⚠️ MCP Server and CLI are BOTH first-class entry points.** Each hosts its own ExcelMcpService instance:
- **MCP Server**: Fully in-process, direct method calls (no pipe)
- **CLI**: Daemon process with named pipe (`excelmcp-cli-{SID}`), sessions persist across CLI invocations
- **Feature parity**: Every action available in MCP must be available in CLI and vice versa
- **Parameter parity**: Same parameters, same defaults, same validation

