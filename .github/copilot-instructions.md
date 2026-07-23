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

**NEVER `--no-verify` (Rule 31):** Never bypass pre-commit hooks with `--no-verify`/`-n`/`HUSKY=0`/`core.hooksPath`. Let all 14 gates run to completion. The `CI Gate` workflow (`.github/workflows/ci.yml`) enforces the **Excel-free** gates on every PR (Release build, COM-leak, success-flag, coverage/naming, MCP-Core, doc-counts, dynamic-cast, plugin-README), but the **Excel-dependent** gates (CLI workflow smoke, MCP smoke) and the packaging deliverables run **only** in pre-commit — so bypassing still ships those unverified. If a hook fails or hangs (incl. environment issues like VBA trust disabled), STOP and ask the user — never bypass.

**PR Review:** Check automated comments immediately (Copilot, GitHub Security). Fix before human review.

**Surgical Testing:** Integration tests take 45+ minutes. ALWAYS test only the feature you changed using `--filter "Feature=<name>"`.

**Test Timeouts:** ALWAYS set an explicit timeout when running tests from terminal or agent tooling so hung Excel/COM runs fail fast instead of blocking the session.

**MCP Parameter Naming:** NEVER use underscores in C# Core interface parameter names. The `McpToolGenerator` calls `StringHelper.ToSnakeCase()` on the C# parameter name to produce the MCP snake_case parameter automatically. Use camelCase in C# that produces the desired snake_case output: `rangeAddress` → `range_address`, `sourceRangeAddress` → `source_range_address`. If the C# name can't produce the desired MCP name via ToSnakeCase, use `[FromString("desiredName")]` attribute instead of underscores in C# names.

**ExcelWriteGuard (Structural Safety):** `Execute()` automatically suppresses `ScreenUpdating` via `ExcelWriteGuard`. Do NOT add `ScreenUpdating` suppression in command code. Calculation suppression is manual and ONLY in value/formula write commands (SetValues, SetFormulas, Append, Write). NEVER suppress `EnableEvents` or `Calculation` universally — Data Model, PivotTable, and Power Query operations depend on them.

**Shutdown Resilience:** ALL workbook close paths (single AND multi-workbook) use `ExcelShutdownService`. Save and Close have retry for transient errors. PID capture has retry for Hwnd=0. `AppDomain.ProcessExit` handler kills tracked Excel PIDs on crash.

**Test Fixture Anti-Pattern:** NEVER use both `IClassFixture<T>` and `[Collection("...")]` with a collection fixture on the same test class. Dual fixtures create concurrent Excel sessions that deadlock during initialization with `maxParallelThreads: 1`. Use ONLY the collection fixture.

**Golden Rule (Diagnose Before Coding):** No changes without a failing test first. Write a test that proves the bug exists, watch it fail, then fix it, then watch it pass. Diagnose root cause before writing any code — spent a full session implementing the wrong fix once (issue was daemon dying, but coded a UI fix + 4 tests, all reverted).

**Bug Fix Pattern Search:** Every bug fix must include a same-pattern search across sibling tools, Core interfaces, generated MCP/CLI surfaces, and tests. Fix matching cases in the same PR or document why each similar case is not affected.

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

## 🔒 Pre-Commit Hooks (15 Automated Gates)

Pre-commit runs `scripts/pre-commit.ps1` which blocks commits if any check fails:

| # | Check | Script | What It Validates |
|---|-------|--------|-------------------|
| 1 | Branch | (inline) | Never commit to `main` directly (Rule 6) |
| 2 | COM Leaks | `check-com-leaks.ps1` | All `dynamic` COM objects have `ComUtilities.Release()` in finally |
| 3 | Coverage + Naming Audit | `audit-core-coverage.ps1` | 100% Core methods exposed via MCP Server and action names stay aligned |
| 4 | MCP-Core Implementation | `check-mcp-core-implementations.ps1` | All enum actions have Core method implementations |
| 5 | Success Flag | `check-success-flag.ps1` | Rule 0: Never `Success=true` with `ErrorMessage` |
| 6 | Release Solution Build | `dotnet build Sbroenne.ExcelMcp.sln -c Release` | Refreshes Release binaries and generated skill outputs used by packaging |
| 6b | Documentation Count Validation | `check-doc-counts.ps1` | All user-facing docs report the code-derived tool/operation counts (26 tools / 234 operations) |
| 7 | CLI Workflow Test | `Test-CliWorkflow.ps1` | E2E CLI workflow smoke test |
| 8 | MCP Smoke Test | `dotnet test --filter "...SmokeTest..."` | All MCP tools functional |
| 9 | CLI Release Deliverables | `dotnet pack` + `dotnet publish` + ZIP | Local CLI NuGet + standalone ZIP match release shapes |
| 10 | MCP Server Release Deliverables | `dotnet pack` + `dotnet publish` + ZIP | Local MCP Server NuGet + standalone ZIP match release shapes |
| 11 | VS Code Extension Package | `npm run package` | Release packaging path succeeds before commit |
| 12 | MCPB Bundle | `mcpb\Build-McpBundle.ps1` | Claude Desktop `.mcpb` bundle builds locally |
| 13 | Agent Skills Deliverables | `Build-AgentSkills.ps1` | Skills ZIP package builds locally |
| 14 | Dynamic Cast Audit | `check-dynamic-casts.ps1` | Every `((dynamic))` cast has a justification comment |

**Install hook:**
```powershell
# From repo root
Copy-Item scripts\pre-commit.ps1 .git\hooks\pre-commit
```

---

## 🧪 LLM Integration Tests (`llm-tests/`)

Separate pytest-based project validating LLM behavior using `pytest-skill-engineering`:

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
- GitHub auth for public-repo operations must use a personal GitHub account (not an EMU account) via `gh auth login --with-token` or a `GITHUB_TOKEN` from that account

**GitHub auth rule for this repo:** When using `gh` against `sbroenne/mcp-server-excel` (issues, PRs, comments, merges), verify the authenticated account is a personal GitHub account. Enterprise Managed User (EMU) accounts have **read-only** (`pull`) access here and cannot create PRs, edit rulesets, disable workflows, delete workflow runs, or access some public-repo API paths.

**Selecting the personal token (admin operations):** The default `GH_TOKEN` may point at the EMU account (`admin:false, push:false`). Check with `gh api repos/sbroenne/mcp-server-excel --jq '.permissions'`. Copilot CLI exposes each authenticated account's token as an env var named `COPILOT_GH_ACCOUNT_github_2E_com_<login>` (e.g. `COPILOT_GH_ACCOUNT_github_2E_com_sbroenne` for the personal owner account). For admin/write work, set `$env:GH_TOKEN = $env:COPILOT_GH_ACCOUNT_github_2E_com_sbroenne` at the start of the command (it does not persist across `powershell` calls, so re-set it each call), then verify `gh api user --jq '.login'` returns `sbroenne`.

**Managing "stale" registered workflows (Actions sidebar):** A workflow whose YAML file was deleted (e.g. `squad-ci.yml`) can linger as an "active" entry; it won't run, and it usually clears itself once its last run ages out. GitHub-managed *dynamic* workflows like `pages-build-deployment` cannot be disabled (`/disable` returns 422) — remove them by deleting all their runs (`DELETE /repos/.../actions/runs/{id}`), after which the sidebar entry disappears. When Pages is already `build_type: workflow`, `pages-build-deployment` will not regenerate.

**Rulesets REST gotchas:** Update a ruleset with `PUT /repos/{owner}/{repo}/rulesets/{id}` (a `PATCH` returns 404) and send the **full** object (`name`, `target`, `enforcement`, `conditions`, `rules`, `bypass_actors`) — a partial body is rejected. Unknown rule parameters are **silently dropped**, so always re-GET and confirm the change persisted. To require automatic Copilot code review, add a rule of type `copilot_code_review` with parameters `review_on_push` and `review_draft_pull_requests` (it is a distinct rule type, NOT a `pull_request` parameter such as the non-existent `automatic_copilot_code_review_enabled`).

**PR/CI enforcement on this repo (single-maintainer):** `sbroenne` is the **sole developer**, so the `main` ruleset requires **0 human approvals** (GitHub forbids approving your own PR — a required-approval rule would deadlock every merge). The effective review gate is the **mandatory `copilot_code_review` rule**. Required status check: **`CI Gate`** (`.github/workflows/ci.yml`), which runs only the **Excel-free** gates because GitHub-hosted runners have **no Excel** (Excel-dependent gates stay local-only in `scripts/pre-commit.ps1`). Merges are squash-only with `delete_branch_on_merge: true` and auto-merge enabled.

**Copilot-review auto-merge gotcha (IMPORTANT):** With the `copilot_code_review` rule active, a PR stays `mergeStateStatus: BLOCKED` — even when every status check is green and no human approval is required — until **Copilot's review comment threads are resolved**. This holds independently of the `pull_request` rule's `required_review_thread_resolution` setting. So the solo auto-merge flow is: address each Copilot comment in code → **reply and resolve the thread** (`resolveReviewThread` GraphQL mutation) → auto-merge then completes on its own. Before resolving, verify any Copilot claim rather than dismissing it (Rule 23); e.g. Copilot's "`contents: read` breaks `actions/upload-artifact`" was a false positive — that action uploads via `ACTIONS_RUNTIME_TOKEN`, not `GITHUB_TOKEN` scopes, and the upload steps concluded `success` with artifacts produced.

**Sister projects (share this repo's conventions):** This repo has two sibling repos under `sbroenne`, both Windows-only C# **MCP Server & CLI** tools that share this repo's source-generator approach, `scripts/pre-commit.ps1` gate model, and the solo PR/CI enforcement described above. They differ in what they automate:
- **`sbroenne/mcp-server-powerpoint`** — automates PowerPoint via the PowerPoint **COM API**, using the same COM-interop layering as this repo (ComInterop → Core → Service → CLI + MCP Server). Cloud-CI constraint mirrors this repo but for **PowerPoint** (GitHub-hosted runners have no Microsoft PowerPoint, so only PowerPoint-free gates run in CI).
- **`sbroenne/mcp-windows`** — the Windows MCP server: drives Windows applications through the **Windows UI Automation (UIA) API** rather than Office COM, so its interop layer differs, but it follows the same MCP Server & CLI structure and conventions. Its cloud-CI constraint is the lack of an interactive desktop/UIA session on GitHub-hosted runners rather than a specific Office app.

When applying fixes/conventions here, consider whether the sibling repos need the same change (and vice versa). Use the same personal-token + auto-merge flow for all three.

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
```powershell
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

