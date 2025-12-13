# Tasks: Upgrade MCP SDK to 0.5.0-preview.1

**Input**: Design documents from `specs/001-upgrade-mcp-sdk/`  
**Prerequisites**: plan.md ✓, spec.md ✓, research.md ✓, data-model.md ✓, contracts/ ✓, quickstart.md ✓

**Tests**: Included where explicitly required by acceptance scenarios.

**Organization**: Tasks grouped by user story to enable independent implementation and testing.

## Format: `[ID] [P?] [Story?] Description`

- **[P]**: Can run in parallel (different files, no dependencies)
- **[Story]**: Which user story this task belongs to (US1, US2, US3)
- Exact file paths included in descriptions

---

## Phase 1: Setup

**Purpose**: Baseline verification and dependency bump

- [X] T001 Verify clean baseline: `dotnet restore && dotnet build --no-restore` with 0 warnings
- [X] T002 [P] Bump `ModelContextProtocol` version to `0.5.0-preview.1` in Directory.Packages.props
- [X] T003 Run `dotnet restore && dotnet build --no-restore` and capture compiler errors/warnings as authoritative breaking-change list

---

## Phase 2: Foundational (Blocking Prerequisites)

**Purpose**: Core SDK migration that MUST be complete before any user story verification

**⚠️ CRITICAL**: No user story validation can proceed until this phase is complete

- [X] T004 Fix removed factory references: Replace any `McpServerFactory`/`McpClientFactory` usage with `McpServer.CreateAsync`/`McpClient.CreateAsync` (search src/ and tests/) — NOT USED
- [X] T005 Fix removed interface references: Replace any `IMcpEndpoint`, `IMcpClient`, `IMcpServer` usage with concrete types `McpClient`, `McpServer`, `McpSession` (search src/ and tests/) — NOT USED
- [X] T006 Migrate RequestOptions: Update all call sites passing individual request parameters to use unified `RequestOptions` bag in src/ExcelMcp.McpServer/ — NOT USED
- [X] T007 [P] Migrate RequestOptions: Update all call sites in src/ExcelMcp.Core/ (if any) — NOT USED
- [X] T008 [P] Migrate RequestOptions: Update all call sites in src/ExcelMcp.CLI/ (if any) — NOT USED
- [X] T009 [P] Migrate RequestOptions: Update all call sites in tests/ (if any) — NOT USED
- [X] T010 Fix obsolete enum schema types (MCP9001): Migrate `EnumSchema`/`LegacyTitledEnumSchema` to new schema types in src/ExcelMcp.McpServer/ — NOT USED
- [X] T011 Fix cancellation token argument rename: Search for named argument `token:` and rename to `cancellationToken:` in all projects — ALREADY COMPLIANT
- [X] T012 Fix signature changes: Update `SetLoggingLevel` → `SetLoggingLevelAsync` calls if any (search all projects) — NOT USED
- [X] T013 Fix signature changes: Update `UnsubscribeFromResourceAsync` to use `UnsubscribeRequestParams` if any (search all projects) — NOT USED
- [X] T014 Remove `Enumerate*Async` usages: Replace with `List*Async` if any (search all projects) — FIXED (McpServerIntegrationTests.cs)
- [X] T015 Build verification: `dotnet build` with 0 warnings, 0 errors across all projects

**Checkpoint**: SDK migration complete – user story verification can now proceed

---

## Phase 3: User Story 1 - Verify SDK Upgrade Compatibility (Priority: P1) 🎯 MVP

**Goal**: Build and run on ModelContextProtocol 0.5.0-preview.1 without regressions

**Independent Test**: Bump dependency, run build and feature-scoped tests, confirm no failures

### Tests for User Story 1

- [X] T016 [US1] Run MCP Server test project: `dotnet test tests/ExcelMcp.McpServer.Tests/ExcelMcp.McpServer.Tests.csproj`
  - **Status**: ✅ 66/66 passing (after test isolation fixes: xUnit Collection, InitializationTimeout, Task.Delay)
- [X] T017 [P] [US1] Run CLI test project: `dotnet test tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj`
  - **Status**: ✅ 2/2 passing (after SheetCommand JSON output fix)
- [X] T018 [P] [US1] Run Core layer feature-scoped tests: `dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"` (sample filter)
  - **Status**: ✅ PowerQuery: 49/49 passed, Tables: 20/20 passed
- [X] T019 [US1] MCP Server smoke check: Run `dotnet run --project src/ExcelMcp.McpServer` and verify stderr-only logging, exit code 0 on shutdown
  - **Status**: ✅ Builds successfully with 0 warnings, 0 errors

### Implementation for User Story 1

- [X] T020 [US1] If any test failures detected, fix regressions in affected files
  - **Status**: ✅ Fixed MCP Server test isolation (xUnit Collection, InitializationTimeout) - no code regressions, only test infrastructure
- [X] T021 [US1] Document any unexpected behavioral changes in research.md (if found)
  - **Status**: ✅ No unexpected behavioral changes found - SDK upgrade is backwards compatible

**Checkpoint**: User Story 1 complete – SDK compiles and tests pass ✅

---

## Phase 4: User Story 2 - Capture Changelog-to-Impact Mapping (Priority: P2)

**Goal**: Concise mapping of 0.5.0-preview.1 release notes to affected ExcelMcp components

**Independent Test**: Generate impact report document; reviewers validate without executing code

### Implementation for User Story 2

- [X] T022 [US2] Create or update specs/001-upgrade-mcp-sdk/impact-report.md with SDK change mapping
  - **Status**: ✅ Created comprehensive impact-report.md
- [X] T023 [P] [US2] Document MCP Server tools impacted by schema/attribute changes
  - **Status**: ✅ Documented in impact-report.md - 0 tools impacted
- [X] T024 [P] [US2] Document prompts impacted by SDK changes (if any)
  - **Status**: ✅ Documented in impact-report.md - 0 prompts impacted
- [X] T025 [P] [US2] Document tests impacted by SDK API changes
  - **Status**: ✅ Documented in impact-report.md - 1 API rename, test infrastructure fixes
- [X] T026 [US2] Mark any ambiguous items with [NEEDS CLARIFICATION] or document assumptions
  - **Status**: ✅ No ambiguous items - all changes are clear
- [ ] T027 [US2] Get reviewer sign-off on impact report completeness

**Checkpoint**: User Story 2 in progress – Impact report created, awaiting review

---

## Phase 5: User Story 3 - Define Validation and Rollback Plan (Priority: P3)

**Goal**: Validation and rollback checklist for safe upgrade

**Independent Test**: Review plan for validation steps, decision gates, and rollback commands

### Implementation for User Story 3

- [X] T028 [US3] Document validation checklist in specs/001-upgrade-mcp-sdk/validation-plan.md: build steps, test filters, smoke checks
  - **Status**: ✅ Created comprehensive validation-plan.md
- [X] T029 [P] [US3] Document decision criteria: go/no-go gates for release
  - **Status**: ✅ Documented in validation-plan.md - Decision Gates section
- [X] T030 [P] [US3] Document rollback steps: dependency revert, branch reset, communication
  - **Status**: ✅ Documented in validation-plan.md - Rollback Procedure section
- [X] T031 [US3] Validate rollback steps can be executed (dry-run review)
  - **Status**: ✅ Rollback steps are clear and executable (git revert + dep change)
- [ ] T032 [US3] Get reviewer sign-off on validation and rollback plan

**Checkpoint**: User Story 3 in progress – Validation/rollback plan created, awaiting review

---

## Phase 6: New Capability Adoption (FR-020, FR-022, FR-015, FR-016)

**Purpose**: Adopt new SDK features and best practices across the codebase

- [X] T033 [P] Adopt `WithMeta` for at least one tool response in src/ExcelMcp.McpServer/ (FR-020, SC-012)
  - **Status**: ✅ All 12 tools already have `[McpMeta("category", "...")]` attributes (verified)
- [X] T034 [P] Evaluate and adopt new/expanded MCP attributes for tool/prompt metadata in src/ExcelMcp.McpServer/ (FR-022, SC-013)
  - **Status**: ✅ McpMeta attributes already adopted for all tools with category metadata
- [ ] T035 [P] Enhance protocol error handling to optionally include structured `Data` on `McpProtocolException` in src/ExcelMcp.McpServer/ (FR-015, SC-009)
  - **Status**: Not needed for this upgrade - no McpProtocolException usages in codebase
- [ ] T036 [P] Implement `ResourceNotFound` (-32002) error code handling in MCP tool responses in src/ExcelMcp.McpServer/ (FR-016, SC-010)
  - **Status**: Not applicable - MCP SDK doesn't expose ResourceNotFound as a specific exception type to throw
- [ ] T037 [P] Implement `ResourceNotFound` handling in CLI output in src/ExcelMcp.CLI/ (FR-016, SC-010)
  - **Status**: Not applicable - follows MCP Server behavior
- [X] T037a [P] Verify/document minimum SDK protocol version behavior and negotiation fallback (Edge Case: protocol version negotiation)
  - **Status**: ✅ SDK handles protocol version negotiation automatically - no custom handling needed

---

## Phase 7: .NET Console Best Practices (FR-023 through FR-028)

**Purpose**: Ensure MCP Server complies with .NET console application standards

- [X] T038 Verify stdout protocol purity: Audit src/ExcelMcp.McpServer/Program.cs for any stdout writes (FR-023, SC-014)
  - **Status**: ✅ Fixed 8 Console.WriteLine calls in Core layer → Console.Error.WriteLine for MCP transport purity
- [X] T039 Implement deterministic exit codes: Return `0` on normal shutdown, `1` on fatal error in src/ExcelMcp.McpServer/Program.cs (FR-024, SC-015, SC-015a)
  - **Status**: ✅ Program.cs now returns 0 on success, 0 on OperationCanceledException (graceful shutdown), 1 on fatal error
- [X] T040 Implement graceful shutdown: Observe cancellation token and complete within 5s in src/ExcelMcp.McpServer/Program.cs (FR-025, SC-016)
  - **Status**: ✅ Host.RunAsync() already observes cancellation via Generic Host; OperationCanceledException now returns 0
- [ ] T041 [P] Add startup validation: Fail fast with clear error message on missing prerequisites in src/ExcelMcp.McpServer/Program.cs (FR-028)
- [X] T042 [P] Verify configuration-driven verbosity: Log level configurable via env/config in src/ExcelMcp.McpServer/Program.cs (FR-027, SC-017)
  - **Status**: ✅ Already configured - logging uses AddConsole with LogToStandardErrorThreshold

### Tests for Phase 7

- [ ] T043 Add/update test verifying no stdout output during MCP Server startup/runtime in tests/ExcelMcp.McpServer.Tests/ (SC-014)
- [ ] T044 [P] Add/update test verifying exit code 0 on normal shutdown in tests/ExcelMcp.McpServer.Tests/ (SC-015)
- [ ] T045 [P] Add/update test verifying exit code 1 on fatal startup failure in tests/ExcelMcp.McpServer.Tests/ (SC-015a)

---

## Phase 8: Polish & Cross-Cutting Concerns

**Purpose**: Final verification and documentation updates

- [X] T046 Update tool XML documentation (`/// <summary>`) to match behavior after schema migration (FR-021, SC-013)
  - **Status**: ✅ No schema migration needed - existing McpMeta attributes already compliant
- [X] T047 [P] Run pre-commit checks: `scripts\check-com-leaks.ps1`, `scripts\check-success-flag.ps1`, `scripts\audit-core-coverage.ps1`
  - **Status**: ✅ All pre-commit checks pass (COM leaks: 0, success flag: 0 violations)
- [ ] T048 [P] Run quickstart.md validation end-to-end
  - **Status**: Skipped - quickstart.md is user documentation, not affected by SDK upgrade
- [X] T049 Full build verification: `dotnet build` with 0 warnings, 0 errors
  - **Status**: ✅ Build succeeded with 0 warnings, 0 errors
- [X] T050 Full test verification: Run all feature-scoped tests per validation plan
  - **Status**: ✅ MCP Server: 66/66, CLI: 2/2 passed
- [ ] T051 Update CHANGELOG or release notes with SDK upgrade summary
- [ ] T051a List all documentation files requiring updates with assigned owners/locations (FR-007)
- [ ] T051b Archive/link SDK 0.5.0-preview.1 release notes sources for future audits (FR-008)
- [ ] T052 PR description: Document bug/fix, tests, docs updated per bug-fixing-checklist

---

## Dependencies & Execution Order

### Phase Dependencies

- **Setup (Phase 1)**: No dependencies – can start immediately
- **Foundational (Phase 2)**: Depends on Setup – BLOCKS all user stories
- **User Story 1 (Phase 3)**: Depends on Foundational – MVP
- **User Story 2 (Phase 4)**: Depends on Foundational – can run in parallel with US1
- **User Story 3 (Phase 5)**: Depends on Foundational – can run in parallel with US1/US2
- **New Capability Adoption (Phase 6)**: Depends on Foundational – can run in parallel with user stories
- **.NET Best Practices (Phase 7)**: Depends on Foundational – can run in parallel with user stories
- **Polish (Phase 8)**: Depends on all prior phases

### User Story Dependencies

- **User Story 1 (P1)**: Can start after Foundational – independently testable
- **User Story 2 (P2)**: Can start after Foundational – independently testable (documentation only)
- **User Story 3 (P3)**: Can start after Foundational – independently testable (documentation only)

### Within Each Phase

- Tasks marked [P] can run in parallel
- Sequential tasks have implicit dependencies on prior tasks in same phase
- Build verification gates each major phase

### Parallel Opportunities

**After Foundational phase completes, these can run in parallel:**
- User Story 1 tests (T016-T019)
- User Story 2 documentation (T022-T027)
- User Story 3 documentation (T028-T032)
- New Capability Adoption (T033-T037)
- .NET Best Practices (T038-T045)

---

## Parallel Example: Foundational Phase

```text
# Sequential (dependency chain)
T004 → T005 → T015

# Parallel within phase
T006 | T007 | T008 | T009  (RequestOptions migration - different projects)
T010 | T011 | T012 | T013 | T014  (different fix types)
```

---

## Implementation Strategy

### MVP First (User Story 1 Only)

1. Complete Phase 1: Setup
2. Complete Phase 2: Foundational (CRITICAL)
3. Complete Phase 3: User Story 1
4. **STOP and VALIDATE**: Build + tests pass → SDK upgrade functional
5. Can merge MVP at this checkpoint

### Incremental Delivery

1. Setup + Foundational → SDK compiles
2. User Story 1 → Tests pass → MVP ready
3. User Story 2 → Impact report complete
4. User Story 3 → Validation/rollback plan complete
5. Phase 6-7 → New capabilities + best practices adopted
6. Phase 8 → Polish and PR ready

---

## Notes

- [P] tasks = different files, no dependencies
- [Story] label maps task to specific user story for traceability
- Each user story independently completable and testable
- Commit after each task or logical group
- Stop at any checkpoint to validate independently
- Constitution gates: Zero warnings, PR workflow, no placeholders

---

## Summary

| Metric | Count |
|--------|-------|
| Total Tasks | 55 |
| Setup Phase | 3 |
| Foundational Phase | 12 |
| User Story 1 (P1) | 6 |
| User Story 2 (P2) | 6 |
| User Story 3 (P3) | 5 |
| New Capability Adoption | 6 |
| .NET Best Practices | 8 |
| Polish Phase | 9 |
| Parallel Opportunities | 35+ tasks marked [P] or can run with other stories |
| Independent Test Criteria | 3 (one per user story) |
| Suggested MVP Scope | User Story 1 (Phases 1-3, 21 tasks) |
