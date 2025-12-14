# Tasks: Status Bar MCP Server Monitor

**Input**: Design documents from `specs/002-statusbar-mcp-status/`
**Prerequisites**: plan.md (required), spec.md (required for user stories), research.md, data-model.md, contracts/

## Phase 1: Setup

- [X] T001: Create feature branch `feature/002-statusbar-mcp-status`
- [X] T002: Confirm VS Code manages the MCP server via the extension's provider; remove any server-spawn logic from plan/tasks
- [X] T003: Add extension configuration for polling interval setting only (no server endpoint settings)
- [X] T004: Scaffold `src/statusBarMcp.ts` and `src/mcpClient.ts`
- [X] T005: Update `package.json` contributes (commands, configuration) and activation events

## Phase 2: Foundational (Blocking Prerequisites)

**Purpose**: Core utilities and MCP access wrappers

- [X] T006: Implement `mcpClient.ts` to obtain a client handle from VS Code's MCP provider and invoke tool calls (no process spawning)
- [X] T007: Add typed contracts for `excel_file` List/Close responses
- [X] T008: Implement polling utility with backoff and cancellation
- [X] T009: Wire activation in `extension.ts` to initialize client and status bar module

## Phase 3: User Story 1 - Quick Server Status Check (Priority: P1) 🎯 MVP

**Goal**: Status bar reflects MCP Server running state; hover shows summary.
**Independent Test**: Start/stop server and observe status bar changes within spec timings.

- [X] T010: Render status bar item "Excel MCP" with server status (Connected/Disconnected)
- [X] T011: Tooltip: show active session count and short guidance
- [X] T012: Poll every ~3s; on error, backoff to 5s and mark Disconnected

## Phase 4: User Story 2 - View Active Sessions (Priority: P2)

**Goal**: Click opens a Quick Pick with active sessions, formatted.
**Independent Test**: Open sessions via MCP, click status bar, confirm list displays correctly.

- [X] T013: On click, fetch sessions via `excel_file` List; display in Quick Pick
- [X] T014: Include per-session details (file path, canClose, isExcelVisible)
- [X] T015: Handle empty state gracefully

## Phase 5: User Story 3 - Close Individual Session (Priority: P3)

**Goal**: Close selected session via Quick Pick.
**Independent Test**: Close one session; verify only that session terminates.

- [X] T016: Provide actions per session: Close, Save & Close
- [X] T017: Call `excel_file` Close with `save: true|false`; respect canClose flag
- [X] T018: Surface JSON errors in notifications; refresh list after success

## Phase 6: User Story 4 - Save & Close Session (Priority: P3)

**Goal**: Save changes then close selected session via Quick Pick.
**Independent Test**: Modify data, Save & Close, reopen to verify persistence.

- [X] T019: Add command palette command: "Excel MCP: Show Sessions" (invokes Quick Pick)
- [X] T020: Add configuration validation and defaults
- [X] T021: Ensure status bar hides when MCP provider/client is unavailable
- [X] T022: Logging and diagnostics (trace errors)
- [X] T023: Unit tests (if applicable) for formatting utilities

## Phase 7: Polish & Documentation

- [X] T024: README updates for extension usage
- [X] T025: Verify Windows-only messaging and Excel closed requirement
- [X] T026: Performance tuning: debounce refresh after actions
- [X] T027: [P] Update `vscode-extension/README.md` with usage, screenshots, and Quickstart references
- [X] T028: Add config setting docs (poll interval) in `vscode-extension/README.md`
- [X] T029: Verify research.md alignment and close any open questions

## Phase 8: Integration Testing (NO MOCKS)

- [X] T030: Implement `McpTestClient` for direct MCP server communication via JSON-RPC stdio
- [X] T031: Add MCP Server Direct tests (initialize, listTools, excel_file actions)
- [X] T032: Add Session Lifecycle tests with real Excel (create, list, close, field validation)
- [X] T033: Add Status Bar visibility tests (hidden initially, visible on connection)
- [X] T034: Verify all tests pass with real MCP server subprocess

---

## Dependencies & Execution Order

- Setup → Foundational → User Stories (US1 → US2 → US3 → US4) → Polish
- Parallel opportunities marked [P]: module creation, package.json contributions, helpers, querying items.

## MVP Scope

- Complete Phase 1–2 and Phase 3 (US1) to deliver MVP: running state indicator + tooltip.

## Totals

- Total tasks: 34
- Completed: 34
- Remaining: 0

## Test Summary

### Integration Tests (Real MCP Server + Real Excel)

| Suite | Tests | Description |
|-------|-------|-------------|
| Status Bar Integration | 7 | Command execution, config, visibility |
| MCP Server Direct | 5 | Initialize, listTools, excel_file actions |
| MCP Session Lifecycle | 4 | Create/list/close sessions with real Excel |

### Unit Tests

| Suite | Tests | Description |
|-------|-------|-------------|
| McpClient | 8 | Polling, error handling, tool calls |
| StatusBarMcp | 9 | State management, visibility, updates |
| ShowSessionsQuickPick | 7 | QuickPick formatting, actions |
| Extension | 5 | Activation, disposal |

**Total: 45 tests (16 integration + 29 unit)**
