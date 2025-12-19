---
description: "Task list for Rename Queries & Data Model Tables feature"
---

# Tasks: Rename Queries & Data Model Tables

**Input**: Design documents from `specs/001-rename-queries-tables/`  
**Feature Branch**: `001-rename-queries-tables`  
**Status**: Implementation ready

**Organization**: Tasks are grouped by user story to enable independent implementation and testing of each story.

## Format: `- [ ] [ID] [P?] [Story?] Description`

- **[P]**: Can run in parallel (different files, no dependencies)
- **[Story]**: Which user story this task belongs to (e.g., [US1], [US2], [US3])
- Exact file paths included in descriptions

---

## Phase 1: Setup (Shared)

 - [X] T001 Align on constraints and behaviors in specs/001-rename-queries-tables/spec.md (trim+CI uniqueness, no-op, case-only rename, never auto-save)
 - [X] T002 Review implementation plan + touchpoints in specs/001-rename-queries-tables/plan.md (Core + MCP + CLI + tests + docs)
 - [X] T003 Review contracts and quickstart examples in specs/001-rename-queries-tables/contracts/rename-contracts.yaml and specs/001-rename-queries-tables/quickstart.md

---

## Phase 2: Foundational (Blocking Prerequisites)

 - [X] T004 Add structured rename result type in src/ExcelMcp.Core/Models/ResultTypes.cs (RenameResult with success/errorMessage/objectType/oldName/newName + normalized fields)
 - [X] T005 Add shared rename name rules helper in src/ExcelMcp.Core/Commands/RenameNameRules.cs (Normalize(trim), empty checks, case-insensitive comparisons, no-op detection)
 - [X] T006 [P] Add new MCP actions to enums in src/ExcelMcp.McpServer/Models/ToolActions.cs (PowerQueryAction.Rename, DataModelAction.RenameTable)
 - [X] T007 [P] Add action string mappings in src/ExcelMcp.McpServer/Models/ActionExtensions.cs for the new enum values (Rule 15: complete mappings)

**Checkpoint**: Core has a shared rename result + shared validation; MCP enums/mappings compile.

---

## Phase 3: User Story 1 - Rename an existing Power Query (Priority: P1) MVP

**Goal**: Rename a workbook Power Query by changing its `Name` via Excel COM, with trim + case-insensitive conflict detection, no-op behavior, and no auto-save.

**Independent Test**: Create/import a query, rename it, verify `List` shows the new name, and `View` content is unchanged.

### Tests for User Story 1

- [X] T008 [P] [US1] Add Power Query rename integration tests in tests/ExcelMcp.Core.Tests/Integration/Commands/PowerQuery/PowerQueryCommandsTests.Rename.cs (success, content unchanged, conflict, missing, invalid, no-op, case-only)

### Implementation for User Story 1

- [X] T009 [US1] Add Rename API to src/ExcelMcp.Core/Commands/PowerQuery/IPowerQueryCommands.cs (Rename(IExcelBatch batch, string oldName, string newName))
- [X] T010 [US1] Implement rename in src/ExcelMcp.Core/Commands/PowerQuery/PowerQueryCommands.Rename.cs (use workbook.Queries; trim+CI collision; no-op skip; case-only attempt; COM release in finally)
- [X] T011 [US1] Wire RenameResult population (objectType=power-query, normalized names) in src/ExcelMcp.Core/Commands/PowerQuery/PowerQueryCommands.Rename.cs

### MCP + CLI surface for User Story 1

- [X] T012 [US1] Add rename routing to src/ExcelMcp.McpServer/Tools/ExcelPowerQueryTool.cs (new action → Core Rename; JSON return for business errors)
- [X] T013 [US1] Add CLI action parsing/output in src/ExcelMcp.CLI/Commands/PowerQuery/PowerQueryCommand.cs (action rename + args --query and --new-name)

### MCP tests for User Story 1

- [X] T014 [US1] Extend MCP smoke coverage in tests/ExcelMcp.McpServer.Tests/Integration/Tools/McpServerSmokeTests.cs to include excel_powerquery rename (create → rename → list)

### Docs for User Story 1

- [X] T015 [US1] Update tool prompt guidance in src/ExcelMcp.McpServer/Prompts/Content/excel_powerquery.md to document rename semantics (trim+CI uniqueness, no-op, case-only allowed)
- [X] T016 [US1] Update tool XML summary in src/ExcelMcp.McpServer/Tools/ExcelPowerQueryTool.cs to mention rename action behavior (no auto-save, no-op, case-only)

**Checkpoint**: US1 complete when Core tests + MCP smoke pass and CLI can execute rename and return RenameResult JSON.

---

## Phase 4: User Story 2 - Rename an existing Data Model table (Priority: P2)

**Goal**: Rename a table in the Excel Data Model (best-effort). Attempt direct COM rename; if blocked, support PQ-backed model tables by renaming the underlying query; otherwise return a clear not-supported business error.

**IMPORTANT FINDING**: ModelTable.Name is READ-ONLY and cannot be changed via the COM API. The implementation correctly detects this limitation and returns a clear error message. Tests verify this behavior.

**Independent Test**: Create a PQ loaded to the Data Model, attempt rename, verify the operation returns a clear error explaining the Excel limitation, and verify the original table name is preserved.

### Tests for User Story 2

- [X] T017 [P] [US2] Add Data Model rename-table integration tests in tests/ExcelMcp.Core.Tests/Integration/Commands/DataModel/DataModelCommandsTests.RenameTable.cs (tests verify Excel limitation is properly detected and reported; rollback preserves original state)

### Implementation for User Story 2

- [X] T018 [US2] Add RenameTable API to src/ExcelMcp.Core/Commands/DataModel/IDataModelCommands.cs (RenameTable(IExcelBatch batch, string oldName, string newName))
- [X] T019 [US2] Implement rename-table in src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.RenameTable.cs (direct COM attempt first; fallback to PQ rename; verify table name updated; rollback on failure; clear error message about immutable table names)
- [X] T020 [US2] Ensure rename-table uses shared rules in src/ExcelMcp.Core/Commands/RenameNameRules.cs (trim+CI uniqueness excluding target; no-op detection)
- [X] T021 [US2] Wire RenameResult population (objectType=data-model-table, normalized names) in src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.RenameTable.cs

### MCP + CLI surface for User Story 2

- [X] T022 [US2] Add rename-table routing to src/ExcelMcp.McpServer/Tools/ExcelDataModelTool.cs (new action → Core RenameTable; JSON return for business errors)
- [X] T023 [US2] Add CLI action parsing/output in src/ExcelMcp.CLI/Commands/DataModel/DataModelCommand.cs (action rename-table + args --table and --new-name)

### MCP tests for User Story 2

- [X] T024 [US2] Extend MCP smoke coverage in tests/ExcelMcp.McpServer.Tests/Integration/Tools/McpServerSmokeTests.cs to include excel_datamodel rename-table (PQ create/load-to-datamodel → rename-table → list-tables)

### Docs for User Story 2

- [X] T025 [US2] Update tool prompt guidance in src/ExcelMcp.McpServer/Prompts/Content/excel_datamodel.md to document rename-table semantics + limitations (best-effort direct rename, PQ-backed fallback, not-supported cases)
- [X] T026 [US2] Update tool XML summary in src/ExcelMcp.McpServer/Tools/ExcelDataModelTool.cs to mention rename-table behavior and constraints (no auto-save, best-effort)

**Checkpoint**: US2 complete when Core tests prove PQ-backed rename works and MCP/CLI surface the action successfully.

---

## Phase 5: User Story 3 - Rename safely within an automated workflow (Priority: P3)

**Goal**: Ensure deterministic behavior and structured results across success + error cases, and ensure MCP returns JSON (not exceptions) for business failures.

**Independent Test**: Exercise rename against missing object, invalid name, conflict, no-op, and case-only scenarios and verify outcomes are deterministic and parseable.

### Tests for User Story 3

- [X] T027 [P] [US3] Add MCP contract tests in tests/ExcelMcp.McpServer.Tests/Integration/Tools/RenameOperationsToolContractTests.cs to verify tool error behavior (business errors return JSON with success=false + isError=true) for excel_powerquery rename and excel_datamodel rename-table

### Implementation for User Story 3

- [X] T028 [US3] Ensure Core operations return consistent RenameResult shapes in src/ExcelMcp.Core/Commands/PowerQuery/PowerQueryCommands.Rename.cs and src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.RenameTable.cs (no-op = success, missing/conflict/invalid = success=false with clear errorMessage)
- [X] T029 [US3] Ensure MCP tools only throw for validation/preconditions and otherwise always serialize Core results in src/ExcelMcp.McpServer/Tools/ExcelPowerQueryTool.cs and src/ExcelMcp.McpServer/Tools/ExcelDataModelTool.cs

**Checkpoint**: US3 complete when deterministic error/success shapes are proven by Core + MCP tests.

---

## Phase 6: Polish & Cross-Cutting Concerns

- [X] T030 [P] Update quickstart to reflect final CLI/MCP parameter names and actions in specs/001-rename-queries-tables/quickstart.md
- [X] T031 [P] Update feature inventory/docs if needed in FEATURES.md (mention rename actions for PowerQuery and DataModel)
- [X] T032 Run Core feature tests in tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj with filter Feature=PowerQuery&RunType!=OnDemand and Feature=DataModel&RunType!=OnDemand
- [X] T033 Run MCP server tests in tests/ExcelMcp.McpServer.Tests/ExcelMcp.McpServer.Tests.csproj (or narrow to smoke/rename contract tests)

---

## Dependencies & Execution Order

### User Story Dependency Graph

- US1 (Power Query rename) → US2 (Data Model rename-table, because PQ-backed fallback relies on rename semantics) → US3 (workflow determinism / contract enforcement)

### Phase Dependencies

- Phase 1 (Setup) blocks nothing but should be completed first.
- Phase 2 (Foundational) blocks all user stories.
- Phase 3 (US1) should complete before US2 because US2’s fallback strategy relies on PQ rename behavior.
- Phase 4 (US2) can start after Phase 2, but may need refactors from US1 if fallback uses shared logic.
- Phase 5 (US3) can be done after US1+US2 implementation exists (tests + tightening behavior).

---

## Parallel Execution Examples

### Parallel Example: Foundational

- T006 Update MCP enums in src/ExcelMcp.McpServer/Models/ToolActions.cs
- T007 Update action mappings in src/ExcelMcp.McpServer/Models/ActionExtensions.cs

### Parallel Example: User Story 1

- T008 Add tests in tests/ExcelMcp.Core.Tests/Integration/Commands/PowerQuery/PowerQueryCommandsTests.Rename.cs
- T015 Update prompt doc in src/ExcelMcp.McpServer/Prompts/Content/excel_powerquery.md

### Parallel Example: User Story 2

- T017 Add tests in tests/ExcelMcp.Core.Tests/Integration/Commands/DataModel/DataModelCommandsTests.RenameTable.cs
- T025 Update prompt doc in src/ExcelMcp.McpServer/Prompts/Content/excel_datamodel.md

---

## Implementation Strategy

### MVP First (US1 Only)

1. Phase 1 → Phase 2
2. Implement + test US1 (Core + MCP + CLI)
3. Validate: US1 acceptance scenarios in specs/001-rename-queries-tables/spec.md

### Incremental Delivery

- Add US2 next (Data Model rename-table, PQ-backed path)
- Add US3 last (contract/determinism hardening + MCP error-shape tests)
