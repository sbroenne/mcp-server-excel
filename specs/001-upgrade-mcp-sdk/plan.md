# Implementation Plan: Upgrade MCP SDK to 0.5.0-preview.1

**Branch**: `001-upgrade-mcp-sdk` | **Date**: 2025-12-13 | **Spec**: `specs/001-upgrade-mcp-sdk/spec.md`
**Input**: Feature specification from `specs/001-upgrade-mcp-sdk/spec.md`

**Note**: This template is filled in by the `/speckit.plan` command. See `.specify/templates/commands/plan.md` for the execution workflow.

## Summary

Upgrade `ModelContextProtocol` to `0.5.0-preview.1`, migrate removed/renamed APIs and schema types (including RequestOptions + schema changes), and validate MCP server runtime contract (JSON-on-business-error, stderr-only logging, deterministic exit codes) with targeted integration tests.

## Technical Context

<!--
  ACTION REQUIRED: Replace the content in this section with the technical details
  for the project. The structure here is presented in advisory capacity to guide
  the iteration process.
-->

**Language/Version**: C# / .NET 8 (TargetFramework `net8.0`)  
**Primary Dependencies**: `ModelContextProtocol` (MCP SDK), `Microsoft.Extensions.Hosting` / `Microsoft.Extensions.Logging`, Application Insights Worker Service (`Microsoft.ApplicationInsights.WorkerService`), CLI uses `Spectre.Console` + `Spectre.Console.Cli`  
**Storage**: N/A (workbook files on disk; Excel COM interop)  
**Testing**: xUnit integration tests (Excel COM). Feature-scoped `dotnet test` filters are the norm.  
**Target Platform**: Windows only (Excel COM interop requirement)  
**Project Type**: Multi-project .NET solution (Core/ComInterop/CLI/MCP Server + tests)  
**Performance Goals**: Fast startup, no stdout noise (MCP transport integrity), resilient shutdown/cleanup, minimal overhead in server loop  
**Constraints**: Zero warnings (`TreatWarningsAsErrors=true`), MCP JSON error contract, COM cleanup via try/finally, no broad try/catch in Core commands, PR workflow  
**Scale/Scope**: Upgrade touches dependency graph + MCP Server entrypoint + tool/schema definitions + selected tests; keep changes surgical and reviewable

## Constitution Check

*GATE: Must pass before Phase 0 research. Re-check after Phase 1 design.*

Gates derived from `.specify/memory/constitution.md` (must hold throughout implementation):

### Pre-Design Gates

- PASS: Success flag integrity (no `Success=true` with `ErrorMessage`).
- PASS: MCP tools return JSON for business errors; `McpException` only for validation/preconditions.
- PASS: Tool descriptions (XML summaries) match behavior and document non-enum parameter conventions.
- PASS: COM cleanup via try/finally + `ComUtilities.Release`, and exception propagation through batch layer.
- PASS: Surgical testing (feature filters), and no save calls unless persistence tests.
- PASS: PR workflow (no commits/pushes without explicit user approval).

GATE STATUS (Pre-Design): PASS (no known required violations for this upgrade).

### Post-Design Gates (Verified after research.md and contracts/)

| Principle | Expected Compliance | Notes |
|-----------|---------------------|-------|
| I. Success Flag | ✅ Maintained | Upgrade does not modify result contract |
| II. JSON Response Contract | ✅ Maintained | WithMeta adoption may enhance metadata, not break contract |
| III. Tool Descriptions | ✅ Maintained + Updated | Adopt new attributes per FR-022 |
| IV. COM Object Lifecycle | ✅ Unaffected | Upgrade is MCP layer only |
| V. Exception Propagation | ✅ Unaffected | Core command layer unchanged |
| VI. COM API First | ✅ Unaffected | |
| VII. Integration-Only Testing | ✅ Applied | Use feature-scoped tests |
| VIII. Test File Isolation | ✅ Applied | |
| IX. Surgical Test Execution | ✅ Applied | `Feature=McpServer` filter |
| X. Save Only for Persistence | ✅ Applied | |
| XI. PR Workflow | ✅ Enforced | Feature branch 001-upgrade-mcp-sdk |
| XII. Test Before Commit | ✅ Enforced | |
| XIII. Never Commit Automatically | ✅ Enforced | |
| XIV. Comprehensive Bug Fixes | N/A | Not a bug fix |
| XV. Check PR Review Comments | ✅ Planned | |
| XVI. Core-MCP Coverage | ✅ Verified | Audit script pre-commit |
| XVII. No Placeholders | ✅ Enforced | |
| XVIII. Trust IDE Warnings | ✅ Applied | Zero warnings post-upgrade |

GATE STATUS (Post-Design): **PASS** – No constitution violations anticipated or required.

## Project Structure

### Documentation (this feature)

```text
specs/001-upgrade-mcp-sdk/
├── plan.md              # This file (/speckit.plan command output)
├── research.md          # Phase 0 output (/speckit.plan command)
├── data-model.md        # Phase 1 output (/speckit.plan command)
├── quickstart.md        # Phase 1 output (/speckit.plan command)
├── contracts/           # Phase 1 output (/speckit.plan command)
└── tasks.md             # Phase 2 output (/speckit.tasks command - NOT created by /speckit.plan)
```

### Source Code (repository root)
```text
src/
├── ExcelMcp.ComInterop/
├── ExcelMcp.Core/
├── ExcelMcp.CLI/
└── ExcelMcp.McpServer/

tests/
├── ExcelMcp.ComInterop.Tests/
├── ExcelMcp.Core.Tests/
├── ExcelMcp.CLI.Tests/
└── ExcelMcp.McpServer.Tests/
```

**Structure Decision**: Multi-project .NET solution; changes will primarily touch `src/ExcelMcp.McpServer` and any shared code paths affected by MCP SDK API changes.

## Complexity Tracking

> **Fill ONLY if Constitution Check has violations that must be justified**

| Violation | Why Needed | Simpler Alternative Rejected Because |
|-----------|------------|-------------------------------------|
| [e.g., 4th project] | [current need] | [why 3 projects insufficient] |
| [e.g., Repository pattern] | [specific problem] | [why direct DB access insufficient] |

No planned constitution violations.

## Phase 0: Outline & Research

Research goals (produce `research.md`):

1. Confirm actual build-breaking deltas when bumping `ModelContextProtocol` to `0.5.0-preview.1` (compile-guided).
2. Identify where SDK schema types are used in MCP server tool definitions and how to migrate away from obsolete enum schema types.
3. Identify where `RequestOptions` is required and update all call sites (MCP Server, Core, CLI, tests).
4. Confirm feasibility and scope of adopting new attributes + `WithMeta` in the current MCP server architecture.
5. Confirm MCP server console best-practice deltas (stdout purity, exit codes, cancellation shutdown) and how to validate them in tests.

## Phase 1: Design & Contracts

Outputs:

- `data-model.md`: “entities” for this change (DependencySet, ImpactReport, ValidationMatrix, RollbackPlan) and their relationships.
- `contracts/`: OpenAPI contract for a small automation surface that can drive upgrade validation (build/test/status). This is documentation-only and used to formalize inputs/outputs for automation.
- `quickstart.md`: Step-by-step local validation: bump package, build, run feature-scoped tests, and run MCP server smoke checks.

## Phase 2: Implementation Planning (for tasks.md)

High-level task breakdown (to be expanded by `/speckit.tasks`):

1. Dependency bump: update central package versions; restore; build.
2. Fix compiler breaks: removed factories/interfaces; RequestOptions migration; signature renames.
3. Schema migration: replace obsolete schema types and update tool/prompt attributes; ensure tool descriptions match behavior.
4. MCP server runtime hardening: stdout purity, deterministic exit codes (fatal=1), graceful cancellation shutdown, configuration-driven verbosity.
5. Update/extend tests: MCP server tests to validate stdout purity + exit codes + selected schema metadata.
6. Verification: build + feature-scoped tests (`Feature=McpServer` etc.); document validation + rollback steps.
