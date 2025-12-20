# Implementation Plan: Rename Queries & Data Model Tables

**Branch**: `001-rename-queries-tables` | **Date**: 2025-12-19 | **Updated**: 2025-12-20 | **Spec**: spec.md

## Summary

Implement rename operations for existing Power Query queries.

For Data Model tables: expose the operation but return a clear error explaining the Excel limitation (table names are immutable).

Approach:
- Power Query: rename by setting the query object `Name` via Excel COM after trim + case-insensitive uniqueness checks.
- Data Model: attempt direct COM rename (will fail), return clear error explaining the limitation and workaround.

## Technical Context

**Language/Version**: C# / .NET 8  
**Primary Dependencies**: Excel COM automation via `dynamic` + `ExcelMcp.ComInterop`, MCP SDK (`ModelContextProtocol`), `System.Text.Json`, CLI via `Spectre.Console.Cli`  
**Storage**: Excel workbook files (`.xlsx`, `.xlsm`)  
**Testing**: xUnit integration tests requiring installed Excel (Core/CLI/MCP server test projects)  
**Target Platform**: Windows with desktop Excel installed
**Project Type**: Multi-project .NET solution (`Sbroenne.ExcelMcp.sln`)  
**Performance Goals**: Typical rename operations complete within 10 seconds for normal workbooks (per spec SC-001)  
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
├── plan.md              # This file
├── research.md          # Phase 0 output (verified 2025-12-20)
├── data-model.md        # Phase 1 output
├── quickstart.md        # Phase 1 output
├── contracts/           # Phase 1 output
└── tasks.md             # Phase 2 output
```

### Source Code (repository root)

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

No constitution violations anticipated.

## Phase 0: Research (Complete - VERIFIED)

Output: `research.md`

Key outcomes:
- Power Query rename: use Excel COM query object `Name` property (with repo-required validation semantics).
- **Data Model table rename: VERIFIED IMPOSSIBLE** via Excel COM API. `ModelTable.Name` is read-only and immutable.

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
   - Attempt direct rename first (will fail with `TargetParameterCountException`).
   - Return clear error message explaining: "Data Model table names are immutable. Delete and recreate with new name."
   - **NO fallback logic** - verified that no workaround exists.

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
- Attempt rename returns clear error about immutable table names.
- Conflict detection for model table rename (checked before COM attempt).
- Missing table returns "not found" error.

### Documentation & Guidance

- Update tool descriptions (XML summaries) for `excel_powerquery` and `excel_datamodel` to include rename semantics and constraints.
- Add/update MCP prompt guidance markdown to document the Data Model limitation.
