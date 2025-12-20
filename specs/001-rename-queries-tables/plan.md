# Implementation Plan: Rename Queries & Data Model Tables

**Branch**: `001-rename-queries-tables` | **Date**: 2025-12-19 | **Spec**: spec.md
**Input**: Feature specification from spec.md

**Note**: This template is filled in by the `/speckit.plan` command. See `.specify/templates/commands/plan.md` for the execution workflow.

## Summary

Implement rename operations for existing Power Query queries and (best-effort) Excel Data Model tables.

Approach:
- Power Query: rename by setting the query object `Name` via Excel COM after trim + case-insensitive uniqueness checks.
- Data Model: attempt direct COM rename (best-effort). If Excel COM does not allow direct rename, support rename for PQ-backed model tables by renaming the underlying query; otherwise return a clear “not supported for this table source” error.

## Technical Context

<!--
  ACTION REQUIRED: Replace the content in this section with the technical details
  for the project. The structure here is presented in advisory capacity to guide
  the iteration process.
-->

**Language/Version**: C# / .NET 8  
**Primary Dependencies**: Excel COM automation via `dynamic` + `ExcelMcp.ComInterop`, MCP SDK (`ModelContextProtocol`), `System.Text.Json`, CLI via `Spectre.Console.Cli`  
**Storage**: Excel workbook files (`.xlsx`, `.xlsm`)  
**Testing**: xUnit integration tests requiring installed Excel (Core/CLI/MCP server test projects)  
**Target Platform**: Windows with desktop Excel installed
**Project Type**: Multi-project .NET solution (`Sbroenne.ExcelMcp.sln`)  
**Performance Goals**: Typical rename operations complete within 10–15 seconds for normal workbooks (per spec SC-001/SC-002)  
**Constraints**: No auto-save, no interactive prompts; COM objects released in `finally`; Core exceptions propagate through batch layer  
**Scale/Scope**: Single-workbook, single-operation commands; batch mode supported for multi-step workflows

## Constitution Check

*GATE: Must pass before Phase 0 research. Re-check after Phase 1 design.*

PASS criteria for this feature:
- Result contract integrity: never return `Success=true` with `ErrorMessage`.
- MCP tools: return JSON for business failures; throw only for validation/preconditions.
- COM lifecycle: acquire COM objects in try; release in finally using `ComUtilities.Release(ref obj!)`.
- Core exception propagation: do not wrap `batch.Execute()` with catch/return patterns.
- Testing discipline: integration tests only; unique file per test; no `Save()` unless testing persistence.
- Coverage enforcement: new Core operations must be surfaced through MCP enums/switches.

## Project Structure

### Documentation (this feature)

```text
specs/001-rename-queries-tables/
├── plan.md              # This file (/speckit.plan command output)
├── research.md          # Phase 0 output (/speckit.plan command)
├── data-model.md        # Phase 1 output (/speckit.plan command)
├── quickstart.md        # Phase 1 output (/speckit.plan command)
├── contracts/           # Phase 1 output (/speckit.plan command)
└── tasks.md             # Phase 2 output (/speckit.tasks command - NOT created by /speckit.plan)
```

### Source Code (repository root)
<!--
  ACTION REQUIRED: Replace the placeholder tree below with the concrete layout
  for this feature. Delete unused options and expand the chosen structure with
  real paths (e.g., apps/admin, packages/something). The delivered plan must
  not include Option labels.
-->

```text
src/
├── ExcelMcp.ComInterop/
├── ExcelMcp.Core/
├── ExcelMcp.McpServer/
└── ExcelMcp.CLI/

tests/
├── ExcelMcp.ComInterop.Tests/
├── ExcelMcp.Core.Tests/
├── ExcelMcp.McpServer.Tests/
└── ExcelMcp.CLI.Tests/
```

**Structure Decision**: Extend existing Core + MCP Server + CLI layers with new rename operations and integration tests.

## Complexity Tracking

> **Fill ONLY if Constitution Check has violations that must be justified**

| Violation | Why Needed | Simpler Alternative Rejected Because |
|-----------|------------|-------------------------------------|

No constitution violations anticipated.

## Phase 0: Research (Complete)

Output: `research.md`

Key outcomes:
- Power Query rename: use Excel COM query object `Name` property (with repo-required validation semantics).
- Data Model table rename: Excel COM object model likely read-only; plan supports best-effort direct rename and PQ-backed rename via query rename.

## Phase 1: Design & Contracts (Complete)

Outputs:
- `data-model.md`: entities and validation rules.
- `contracts/rename-contracts.yaml`: conceptual request/response shapes for CLI/MCP.
- `quickstart.md`: CLI + MCP usage.

Post-design constitution check: still PASS (no new violations introduced by design).

## Phase 2: Implementation Plan

### Core (src/ExcelMcp.Core)

1) Add rename operations:
  - Power Query: add `Rename(IExcelBatch batch, string oldName, string newName)` to `IPowerQueryCommands` + implement in `PowerQueryCommands`.
  - Data Model: add `RenameTable(IExcelBatch batch, string oldName, string newName)` to `IDataModelCommands` + implement in `DataModelCommands`.

2) Validation behavior (both operations):
  - Normalize by trimming input names.
  - Reject empty/whitespace new names.
  - No-op success when normalized names are equal.
  - Uniqueness check is case-insensitive within the target scope, excluding the target being renamed.
  - Case-only rename allowed: attempt COM rename.

3) COM implementation details:
  - Acquire COM objects (e.g., `workbook.Queries`, query, model, model tables) in try block.
  - Release every COM object in finally via `ComUtilities.Release(ref obj!)`.
  - Do not wrap errors with Core-level catch blocks; let exceptions flow to batch layer.

4) Data Model rename strategy:
  - Attempt direct rename first (best-effort) if late-binding permits.
  - If direct rename is blocked (read-only / COM exception), attempt rename for PQ-backed tables by renaming the underlying query and verifying the new table name appears in `ListTables`.
  - Otherwise return a structured failure result stating the table source cannot be renamed.

### MCP Server (src/ExcelMcp.McpServer)

1) Add new actions to the relevant tool enums and switch routing:
  - `excel_powerquery`: add `rename` action.
  - `excel_datamodel`: add `rename-table` action.

2) Follow MCP result contract:
  - Validation/precondition failures throw `McpException`.
  - Business errors return serialized JSON result with `success: false`.

### CLI (src/ExcelMcp.CLI)

1) Power Query CLI:
  - Add `rename` action to `PowerQueryCommand`.
  - Add settings `--query` (old) and `--new-name` (new).
  - Output JSON result.

2) Data Model CLI:
  - Add `rename-table` action to `DataModelCommand`.
  - Add settings `--table` (old) and `--new-name` (new).

### Tests (tests/*)

Add integration tests (feature-scoped):

PowerQuery:
- Rename existing query succeeds.
- Rename conflicts are rejected (case-insensitive + trim).
- Rename missing query fails.
- Trim-equal rename is no-op success.
- Case-only rename attempts COM rename (assert final list contains new casing).

DataModel:
- Create a PQ loaded to Data Model, then rename the PQ-backed model table and verify list tables reflects new name.
- Conflict detection for model table rename.

### Documentation & Guidance

- Update tool descriptions (XML summaries) for `excel_powerquery` and `excel_datamodel` to include rename semantics and constraints.
- Add/update MCP prompt guidance markdown if a tool-specific prompt exists for these tools.
