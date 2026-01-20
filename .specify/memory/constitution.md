<!--
=== SYNC IMPACT REPORT ===
Version change: 1.0.0 → 1.1.0
Modified sections:
  - Technical Constraints → Platform Requirements: .NET 8 → .NET 10
Templates requiring updates:
  - plan-template.md: ✅ Compatible (no .NET version reference)
  - spec-template.md: ✅ Compatible (technology-agnostic)
  - tasks-template.md: ✅ Compatible (no .NET version reference)
Follow-up TODOs:
  - global.json: Update SDK version to .NET 10 SDK
  - README.md: Update .NET badge from 8.0 to 10
  - Directory.Build.props: Update TargetFramework to net10.0
  - GitHub workflows: Update .NET version in CI/CD
Bump rationale: MINOR - Technical constraint updated (target framework)
===========================
-->

# ExcelMcp Constitution

> **18 principles derived from critical-rules.instructions.md and architecture documentation**

## Core Principles

### Category A: Result Contract Integrity

#### I. Success Flag Must Match Reality (NON-NEGOTIABLE)

**NEVER set `Success = true` when `ErrorMessage` is set.**

- **Invariant**: `Success == true` ⟹ `ErrorMessage == null || ErrorMessage == ""`
- Set `Success = true` only in try blocks on actual success
- Always set `Success = false` in catch blocks before setting `ErrorMessage`
- LLMs and MCP clients rely on this flag for workflow decisions

**Rationale**: LLMs see `Success=true` and assume operations succeeded. 43 violations found in 2025-01-28 audit caused silent failures and data corruption.

#### II. MCP Server JSON Response Contract

**MCP tools MUST return JSON responses for business errors, NOT throw exceptions.**

- Parameter validation errors: throw `McpException` (protocol error)
- Business logic errors (table not found, query failed): return JSON with `Success = false`
- Always serialize Core Command results directly
- Clients expect `{"success": false, "errorMessage": "..."}` format with HTTP 200

**Rationale**: MCP specification defines two error mechanisms: protocol errors (exceptions) and tool execution errors (JSON responses).

#### III. Tool Descriptions Must Match Behavior

**Tool XML documentation (`/// <summary>`) is extracted by MCP SDK and sent to LLMs. It must be accurate.**

- Update descriptions when changing defaults or behavior
- Document non-enum parameter values (loadDestination, formatCode, etc.)
- Do NOT include enum action lists (SDK auto-generates)

**Rationale**: LLMs use tool descriptions for server-specific guidance. Outdated descriptions cause incorrect workflow assumptions.

### Category B: COM Object Management

#### IV. COM Object Lifecycle with Finally Blocks (NON-NEGOTIABLE)

**All COM objects MUST be released in finally blocks using try-finally pattern.**

- Declare COM objects as `dynamic?` nullable before try block
- Acquire COM objects in try block
- Release in finally block with null checks via `ComUtilities.Release(ref obj!)`
- **NEVER** use catch blocks to swallow exceptions
- **NEVER** suppress exceptions with error result returns

**Rationale**: COM objects leak if Release() is not reached before exception. Finally blocks execute regardless of exception state.

#### V. Exception Propagation Through Batch Layer (NON-NEGOTIABLE)

**Core Commands: Let exceptions propagate naturally through `batch.Execute()`.**

- `batch.Execute()` catches exceptions via TaskCompletionSource
- Double-wrapping (catch + return error result) loses stack context
- **Allowed**: Loop continuations, optional property access, specific error routing
- **Forbidden**: `catch (Exception ex) { return new Result { Success = false, ErrorMessage = ex.Message }; }`

**Rationale**: Exception handling belongs at the batch layer. Pattern removed from 200+ methods in Nov 2025.

#### VI. COM API First

**Use Excel COM API for everything it supports.**

- Only use external libraries (TOM) for features Excel COM doesn't provide
- Validate against [Microsoft Excel VBA docs](https://learn.microsoft.com/office/vba/api/overview/excel) before adding dependencies
- Excel collections use 1-based indexing, NOT 0-based
- Search [NetOffice repo](https://github.com/NetOfficeFw/NetOffice) for working examples before implementing
- See also: Technical Constraints → Platform Requirements

**Rationale**: Excel COM is quirky. Real-world examples prevent common pitfalls.

### Category C: Testing Discipline

#### VII. Integration-Only Testing

**No unit tests. All tests are integration tests against real Excel instances.**

- Excel COM cannot be meaningfully mocked (dynamic COM objects)
- Integration tests ARE our unit tests—verify business logic through COM
- Tests MUST verify actual Excel state (round-trip validation)
- See `docs/ADR-001-NO-UNIT-TESTS.md` for full rationale

**Rationale**: Testing mocked COM objects proves nothing about real Excel behavior.

#### VIII. Test File Isolation

**Each test creates unique file. NEVER share test files between tests.**

- Use `CoreTestHelper.CreateUniqueTestFile()` for every test
- VBA tests use `.xlsm` extension (NOT .xlsx renamed)
- Binary assertions only (NO "accept both" patterns)
- All required traits present (Category, Speed, Layer, RequiresExcel, Feature)

**Rationale**: Shared files cause test pollution, file lock issues, and maintenance nightmares.

#### IX. Surgical Test Execution

**Run tests ONLY for the specific code you modified.**

- Use `--filter "Feature=<feature>&RunType!=OnDemand"` for feature-specific tests
- Full test suite (45+ minutes) runs in CI/CD only
- Debug test failures one by one, never all tests at once

**Rationale**: Integration tests require Excel COM and are SLOW. Running all tests wastes time.

#### X. Save Only for Persistence Tests

**Tests must NOT call `batch.Save()` unless explicitly testing persistence.**

- FORBIDDEN: Tests only verifying operation success or in-memory state
- REQUIRED: Round-trip tests verifying data persists after close/reopen
- Save is slow (~2-5s). Removing unnecessary saves makes tests 50%+ faster

**Rationale**: Save operations slow down test suites significantly.

### Category D: Development Workflow

#### XI. Pull Request Workflow (NON-NEGOTIABLE)

**All changes MUST go through Pull Requests. Direct commits to main are prohibited.**

- Create feature branch → Make changes → Push → Create PR → CI/CD + review → Merge
- Pre-commit hooks block commits to main
- Check and fix all automated PR review comments before human review

**Rationale**: PR workflow ensures code review, CI/CD validation, and version management.

#### XII. Test Before Commit (NON-NEGOTIABLE)

**NEVER commit, push, or create PRs without first running tests for changed code.**

- Build must pass (0 warnings, 0 errors)
- Run relevant tests with feature filter
- Pre-commit hooks must pass (COM leaks, success flag, coverage)
- Document test results in commit message

**Rationale**: Prevents breaking changes from reaching main, wastes team time debugging failures.

#### XIII. Never Commit Automatically (NON-NEGOTIABLE)

**All commits, pushes, and merges must require explicit user approval.**

- No automated tools may commit without user confirmation
- No background or silent commits allowed
- Agents must prompt before any repository modification

**Rationale**: Prevents accidental changes, enforces review, ensures user control.

#### XIV. Comprehensive Bug Fixes

**Every bug fix MUST include all 6 components.**

1. Code Fix: Minimal surgical changes
2. Tests: 5-8 new tests (regression + edge cases)
3. Documentation: Update 3+ files
4. Workflow Hints: Update SuggestedNextActions
5. Quality Verification: Build passes, tests green
6. PR Description: Comprehensive summary

**Rationale**: Incomplete bug fixes lead to regressions and confusion.

#### XV. Check PR Review Comments

**After creating PR, check and fix all automated review comments immediately.**

- Common reviewers: Copilot, GitHub Advanced Security
- Fix issues: inheritdoc, null checks, functional style, security warnings
- Request human review only after all automated issues resolved

**Rationale**: Automated reviewers catch code quality issues early.

### Category E: Code Quality

#### XVI. Core-MCP Coverage Enforcement

**Every Core Commands method MUST be exposed via MCP Server enum-based routing.**

- Compiler (CS8524) enforces enum coverage in switch expressions
- Pre-commit hook runs `audit-core-coverage.ps1`
- 8-step workflow: Interface → Implementation → Enum → ToActionString → Switch → Method → Build → Docs

**Rationale**: Prevents dead code and ensures all capabilities accessible to AI assistants.

#### XVII. No Placeholders or TODO Markers

**Code must be complete before commit.**

- No TODO, FIXME, HACK, or XXX markers in source code
- No `NotImplementedException`—full implementation only
- Delete commented-out code (use git history)
- Exception: Documentation files only

**Rationale**: Placeholders accumulate and become permanent. Pre-commit hook blocks.

#### XVIII. Trust IDE Warnings (NON-NEGOTIABLE)

**When VS Code, linters, or tooling shows errors, TRUST THEM.**

- Assume warnings are CORRECT until proven otherwise
- To disprove: Run code and verify OR find authoritative documentation
- "I think it will work" is NOT verification
- If uncertain, use simpler approach that doesn't trigger warnings

**Rationale**: Dismissing valid warnings leads to broken code reaching production.

## Technical Constraints

### Platform Requirements

- **Windows Only**: COM interop is Windows-specific
- **Excel Required**: Microsoft Excel 2016+ must be installed
- **Desktop Environment**: Controls actual Excel process (not for server-side processing)
- **.NET 10 SDK**: Required for build and development

### Architecture Layers

1. **ComInterop** (`src/ExcelMcp.ComInterop`): Reusable COM automation patterns (STA threading, session management, batch operations)
2. **Core** (`src/ExcelMcp.Core`): Excel-specific business logic (Power Query, VBA, worksheets, etc.)
3. **CLI** (`src/ExcelMcp.CLI`): Command-line interface for scripting
4. **MCP Server** (`src/ExcelMcp.McpServer`): Model Context Protocol for AI assistants

### Build Quality

- `TreatWarningsAsErrors=true`: Zero warnings policy
- Security analyzers enabled (CA2100, CA3003, CA3006, CA5389, CA5390, CA5394 are errors)
- See Principle XVII for TODO/FIXME marker rules

## Development Workflow

### Test Execution

```powershell
# Development (fast feedback, excludes VBA)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Session/batch code changes (MANDATORY)
dotnet test --filter "RunType=OnDemand"

# Feature-specific (surgical testing)
dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"
```

### Pre-Commit Requirements

1. Tests pass for code you changed
2. Build passes with 0 warnings
3. `scripts\check-com-leaks.ps1` reports 0 leaks
4. `scripts\check-success-flag.ps1` passes
5. `scripts\audit-core-coverage.ps1` shows 100% coverage

### Bug Fix Completeness

See **Principle XIV** for the 6 required components. Detailed checklist: `.github/instructions/bug-fixing-checklist.instructions.md`

### Release Process

- **Version Tags**: `v1.2.3` releases ALL components (MCP Server, CLI, VS Code Extension, MCPB)
- **Semantic Versioning**: MAJOR (breaking), MINOR (features), PATCH (fixes)
- Versions auto-managed by unified release workflow—never update manually

## Governance

- **Constitution supersedes all other practices**: When in conflict, this document wins
- **Amendments require**: Documentation update, PR review, migration plan if breaking
- **Compliance verification**: All PRs must verify adherence to principles
- **Complexity justification**: Deviations from simplicity require documented rationale
- **Runtime guidance**: See `.github/copilot-instructions.md` and `.github/instructions/` for implementation details

**Version**: 1.1.0 | **Ratified**: 2025-12-09 | **Last Amended**: 2025-12-28
