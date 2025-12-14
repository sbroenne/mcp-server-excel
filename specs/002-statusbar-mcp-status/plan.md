# Implementation Plan: Status Bar MCP Server Monitor

**Branch**: `002-statusbar-mcp-status` | **Date**: 2025-12-14 | **Spec**: specs/002-statusbar-mcp-status/spec.md
**Input**: Feature specification from `specs/002-statusbar-mcp-status/spec.md`
**Status**: ✅ COMPLETE

## Summary

Add a persistent VS Code status bar item labeled "Excel MCP" that reflects MCP Server running state and, on click, opens a Quick Pick to list active sessions with actions to Close or Save & Close. Status bar is hidden until MCP server responds successfully (no disconnected/initializing states shown). Status updates occur within 5 seconds; actions complete within expected bounds.

## Implementation Details

### Files Created/Modified

| File | Purpose |
|------|---------|
| `src/statusBarMcp.ts` | Status bar UI component with visibility control |
| `src/mcpClient.ts` | MCP client wrapper for tool calls via vscode.lm.tools |
| `src/showSessionsQuickPick.ts` | Quick Pick UI for session list and actions |
| `src/extension.ts` | Activation wiring |
| `tests/suite/statusBar.test.ts` | Status bar integration tests |
| `tests/suite/mcpServerDirect.test.ts` | MCP server + session lifecycle tests |
| `tests/suite/mcpTestClient.ts` | Test helper for direct MCP server communication |
| `tests/unit/*.test.ts` | Unit tests for all components |

### Key Design Decisions

1. **Status Bar Hidden by Default**: `isVisible` property tracks visibility; bar only shown after first successful poll
2. **No Mocks in Integration Tests**: `McpTestClient` spawns real MCP server subprocess via JSON-RPC stdio
3. **Session Creation Pattern**: CreateEmpty → Open (CreateEmpty doesn't return sessionId)
4. **Polling**: Configurable interval (default 3000ms) with backoff on errors

## Technical Context

<!--
  ACTION REQUIRED: Replace the content in this section with the technical details
  for the project. The structure here is presented in advisory capacity to guide
  the iteration process.
-->

**Language/Version**: TypeScript (VS Code Extension), .NET 8 (MCP Server)  
**Primary Dependencies**: VS Code Extension API (StatusBarItem, QuickPick), MCP Server JSON tools (`excel_file`, `excel_worksheet`, etc.)  
**Storage**: N/A  
**Testing**: VS Code extension integration tests (manual), MCP Server integration tests (see docs/testing), NEEDS CLARIFICATION for automated UI tests  
**Target Platform**: Windows (Excel COM), VS Code Desktop  
**Project Type**: VS Code extension feature + MCP server query endpoints  
**Performance Goals**: Status updates within 5s, hover tooltip within 500ms, actions within 3-5s per spec  
**Constraints**: COM interop exclusivity (Excel must be closed for MCP COM ops per instructions), JSON-only business error handling (MCP), finally-based COM cleanup  
**Scale/Scope**: Up to 10 concurrent sessions displayed; more should degrade gracefully

## Constitution Check

*GATE: Must pass before Phase 0 research. Re-check after Phase 1 design.*

- Result Contract Integrity: PASS (UI consumes MCP JSON success flags; no exceptions for business errors)
- COM Lifecycle: PASS (No new COM code in extension; MCP Server already enforces try/finally release)
- Testing Discipline: PASS (feature uses integration patterns; VS Code UI tested via manual steps; server-side remains covered)
- Development Workflow: PASS (feature branch in place; PR-only; tests before commit)
- Code Quality: PASS (No TODOs; enum mappings unaffected; tool descriptions may need update if new endpoints are added)

## Project Structure

### Documentation (this feature)

```text
specs/[###-feature]/
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
# [REMOVE IF UNUSED] Option 1: Single project (DEFAULT)
src/
├── models/
├── services/
├── cli/
└── lib/

tests/
├── contract/
├── integration/
└── unit/

# [REMOVE IF UNUSED] Option 2: Web application (when "frontend" + "backend" detected)
backend/
├── src/
│   ├── models/
│   ├── services/
│   └── api/
└── tests/

frontend/
├── src/
│   ├── components/
│   ├── pages/
│   └── services/
└── tests/

# [REMOVE IF UNUSED] Option 3: Mobile + API (when "iOS/Android" detected)
api/
└── [same as backend above]

ios/ or android/
└── [platform-specific structure: feature modules, UI flows, platform tests]
```

**Structure Decision**: Single repository with VS Code extension changes under `vscode-extension/src/` (new `statusBarMcp.ts` module, wiring in `extension.ts`), and optional MCP Server additions under `src/ExcelMcp.McpServer/Tools/ExcelTools.cs` (read-only session list endpoint if needed). No changes to Core required.

## Complexity Tracking

> **Fill ONLY if Constitution Check has violations that must be justified**

| Violation | Why Needed | Simpler Alternative Rejected Because |
|-----------|------------|-------------------------------------|
| [e.g., 4th project] | [current need] | [why 3 projects insufficient] |
| [e.g., Repository pattern] | [specific problem] | [why direct DB access insufficient] |
