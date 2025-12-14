# Feature Specification: Status Bar MCP Server Monitor

**Feature Branch**: `002-statusbar-mcp-status`  
**Created**: December 14, 2025  
**Status**: Draft  
**Input**: User description: "Add status bar functionality to VSCode Extension showing MCP Server running status and active sessions with hover actions to close or save & close sessions"

## Implementation Status

**Status**: ✅ COMPLETE (December 14, 2025)

### Summary
All 4 user stories implemented. Status bar shows MCP connection status with session count. Click opens Quick Pick with active sessions and Close/Save & Close actions. Status bar is hidden until MCP server successfully responds (no "Disconnected" or "Initializing" states shown).

### Key Design Decisions
1. **Status Bar Visibility**: Hidden until first successful MCP poll; hidden when disconnected (clean UX)
2. **No Mocks in Tests**: All integration tests use real MCP server subprocess + real Excel
3. **Session Creation**: Uses CreateEmpty then Open pattern (CreateEmpty doesn't return sessionId)

## Clarifications

### Session 2025-12-14
- Q: What interaction model should be used for listing sessions and invoking Close / Save & Close actions? → A: Click opens a Quick Pick listing sessions; actions selected via Quick Pick.
- Q: Should status bar show "Disconnected" or "Initializing" states? → A: No, status bar should be completely hidden until MCP server is connected. Never show disconnected state.

## User Scenarios & Testing *(mandatory)*

### User Story 1 - Quick Server Status Check (Priority: P1)

As a developer using the Excel MCP extension, I want to see at a glance whether the MCP Server is running so I can verify my environment is ready without checking logs or running commands.

**Why this priority**: This is the core value proposition - immediate visibility of server status is the foundation of all monitoring functionality.

**Independent Test**: Can be fully tested by starting/stopping the MCP server and observing the status bar indicator changes, delivering immediate value even without hover interactions.

**Acceptance Scenarios**:

1. **Given** the VSCode Extension is installed and MCP Server is not running, **When** I look at the status bar, **Then** I see an indicator showing "Excel MCP" (or icon) with a visual state indicating "not running"
2. **Given** the MCP Server starts successfully, **When** the extension detects the running server, **Then** the status bar indicator updates to show "running" state within 2 seconds
3. **Given** the MCP Server is running, **When** the server stops or crashes, **Then** the status bar indicator updates to show "not running" state within 5 seconds

---

### User Story 2 - View Active Sessions (Priority: P2)

As a developer working with multiple Excel files, I want to hover over the status bar item to see all active MCP sessions so I can understand which files currently have open connections.

**Why this priority**: Session visibility is critical for understanding resource usage and identifying stale sessions, but the feature is still useful without this capability.

**Independent Test**: Can be fully tested by opening Excel files through MCP operations, hovering over the status bar to see a read-only tooltip, and clicking the status bar to open a Quick Pick presenting sessions and actions.

**Acceptance Scenarios**:

1. **Given** the MCP Server is running with no active sessions, **When** I hover over the status bar indicator, **Then** I see a tooltip showing "No active sessions"
2. **Given** the MCP Server has 2 active sessions with files "Budget.xlsx" and "Sales.xlsx", **When** I hover over the status bar, **Then** I see a list showing both file names in a clear, readable format
3. **Given** I have 5+ active sessions, **When** I hover over the status bar, **Then** the session list is displayed in a scrollable format with the most recent sessions shown first
4. **Given** a session has been idle for 30+ minutes, **When** I view the session list, **Then** the session shows an indicator of inactivity (e.g., time since last operation)

---

### User Story 3 - Close Individual Session (Priority: P3)

As a developer who wants to clean up resources, I want to close individual MCP sessions from the status bar hover menu so I can free up Excel connections without closing VSCode or running terminal commands.

**Why this priority**: Enhances user control and resource management, but core monitoring works without it.

**Independent Test**: Can be fully tested by creating sessions, hovering over status bar, clicking a close action for one session, and verifying that session is terminated while others remain active.

**Acceptance Scenarios**:

1. **Given** the hover menu shows 2 active sessions, **When** I click the "Close" action for one session, **Then** that session is terminated immediately and removed from the list
2. **Given** I attempt to close a session that has unsaved changes, **When** I click "Close", **Then** the session closes without saving (user acknowledges data loss risk)
3. **Given** a close operation fails due to a locked Excel file, **When** the close completes, **Then** I see an error message indicating the failure reason

---

### User Story 4 - Save and Close Session (Priority: P3)

As a developer who wants to preserve work, I want to save and close individual MCP sessions from the status bar hover menu so I can persist changes before terminating the connection.

**Why this priority**: Provides data safety, but users can manually save before closing as a workaround.

**Independent Test**: Can be fully tested by creating a session with modified data, using "Save & Close" action, reopening the file, and verifying changes persisted.

**Acceptance Scenarios**:

1. **Given** the hover menu shows a session with a modified Excel file, **When** I click "Save & Close", **Then** the file is saved to disk and the session is terminated
2. **Given** a save operation fails due to file permissions, **When** I click "Save & Close", **Then** I see an error message and the session remains open
3. **Given** I click "Save & Close" on multiple sessions rapidly, **When** the operations execute, **Then** each session completes its save and close independently without interfering with others

---

### Edge Cases

- What happens when the MCP Server crashes while sessions are active? (Status bar should show "not running" and session list becomes unavailable)
- How does the system handle sessions that become orphaned (server running but Excel process terminated)? (Session list should show these with a warning indicator)
- What happens if VSCode loses connection to the MCP Server temporarily? (Status bar should show "connection lost" state rather than "not running" to differentiate)
- How does the extension behave when multiple VSCode windows are open with the same extension? (Each window should independently track and display server status)
- What happens when hovering over the status bar while a close/save operation is in progress? (Session should show "closing..." or "saving..." state)
- How does the system handle extremely long file paths or file names in the session list? (Truncate with ellipsis and show full path in tooltip)

## Requirements *(mandatory)*

### Functional Requirements

- **FR-001**: Extension MUST display a status bar item labeled "Excel MCP" (or icon) that is visible at all times
- **FR-002**: Status bar item MUST visually indicate whether the MCP Server is running (different colors, icons, or text states)
- **FR-003**: Extension MUST detect MCP Server state changes within 5 seconds of the change occurring
- **FR-004**: Users MUST be able to hover over the status bar item to reveal a read-only tooltip (summary: server state and session count)
- **FR-005**: Clicking the status bar MUST open a Quick Pick listing active sessions when the server is running
- **FR-006**: Session list MUST display file names or identifiable information for each active session
- **FR-007**: The Quick Pick MUST provide per-session actions including "Close"
- **FR-008**: The Quick Pick MUST provide per-session actions including "Save & Close"
- **FR-009**: Close actions MUST terminate the specified session without affecting other sessions
- **FR-010**: Save & Close actions MUST persist file changes before terminating the session
- **FR-011**: Extension MUST handle server connection failures gracefully without crashing
- **FR-012**: Session information MUST be displayed in a user-friendly format (not raw IDs or internal codes)
- **FR-013**: Extension MUST update session list in real-time as sessions are created or destroyed
- **FR-014**: Status bar item MUST be clickable to open the Quick Pick session/action UI (default action)

### Key Entities

- **MCP Server**: The running Excel MCP server process that the extension monitors; has states (running, stopped, crashed, unreachable)
- **MCP Session**: An active connection between the MCP Server and an Excel file; has attributes (file path, session ID, creation time, last activity time, modification status)
- **Status Bar Item**: The visual indicator in VSCode's status bar; has states (running, not running, warning) and displays session count
- **Hover Menu**: The contextual information panel that appears when hovering over the status bar item; contains session list and action buttons

## Success Criteria *(mandatory)*

### Measurable Outcomes

- **SC-001**: Users can determine if MCP Server is running within 1 second of looking at the VSCode window
- **SC-002**: Session information becomes visible within 500ms of hovering over the status bar item
- **SC-003**: Close operations complete within 3 seconds under normal conditions
- **SC-004**: Save & Close operations complete within 5 seconds for files under 50MB
- **SC-005**: Status bar updates reflect server state changes within 5 seconds with 99% accuracy
- **SC-006**: Users can successfully close sessions without terminal commands in 100% of non-error cases
- **SC-007**: Hover menu remains responsive and readable with up to 10 active sessions
- **SC-008**: Extension handles server connection loss without requiring VSCode restart in 100% of cases

## Assumptions *(optional)*

- MCP Server exposes an API or mechanism to query running state and active sessions
- Extension has permissions to execute close and save operations on behalf of the user
- VSCode extension can poll or subscribe to server state changes
- Session information includes file path or identifiable metadata
- Users have necessary file system permissions for save operations
- Standard VSCode status bar API supports the required interaction patterns
- MCP Server can differentiate between graceful close and forced termination

## Out of Scope *(optional)*

- Starting/stopping the MCP Server from the status bar (separate feature - server lifecycle management)
- Editing Excel file content directly from VSCode
- Bulk operations (close all, save all) - can be added later based on user feedback
- Session history or logging (tracking previously closed sessions)
- Performance metrics per session (CPU, memory usage)
- File preview or thumbnails in the hover menu
- Keyboard shortcuts for status bar interactions
- Customization of status bar item appearance (color schemes, icon choices)
- Advanced filtering or search within session list

## Dependencies *(optional)*

- MCP Server must provide an API endpoint or IPC mechanism to query server status
- MCP Server must expose session enumeration capability
- MCP Server must support remote session close operations
- MCP Server must support save operations before closing sessions
- VSCode Extension API must support status bar items with hover interactions
- Extension must maintain an active connection or polling mechanism to the MCP Server

## Integration Tests *(implemented)*

### Test Architecture
All tests use **real MCP server** and **real Excel** - NO MOCKS. The `McpTestClient` class spawns the MCP server as a subprocess and communicates via JSON-RPC 2.0 over stdio.

### Status Bar Integration Tests (`tests/suite/statusBar.test.ts`)

| Test | Description | Type |
|------|-------------|------|
| showSessions command should execute without error | Executes the command palette command | Integration |
| configuration changes should be reflected | Verifies pollIntervalMs config bounds (1000-60000) | Integration |
| extension should handle rapid command execution | Tests 3x rapid showSessions calls | Integration |
| status bar should only show when MCP server is connected | Verifies hidden initially, hidden before first poll | Integration |
| status bar becomes visible when MCP server responds successfully | Uses real MCP server to test visibility toggle | Integration |
| command palette should show Excel MCP command | Verifies command registration | Integration |
| configuration should have correct type constraints | Type checks pollIntervalMs | Integration |

### MCP Server Direct Integration Tests (`tests/suite/mcpServerDirect.test.ts`)

| Test | Description | Type |
|------|-------------|------|
| server should respond to initialize request | JSON-RPC initialize handshake | Integration |
| server should list available tools | Verifies excel_file, excel_worksheet, etc. exist | Integration |
| excel_file List action should return empty sessions | Tests List action with no sessions | Integration |
| excel_file Test action should check Excel availability | Tests Test action | Integration |
| tool schema should have proper input definitions | Validates inputSchema structure | Integration |

### MCP Session Lifecycle Tests (`tests/suite/mcpServerDirect.test.ts`)

| Test | Description | Type |
|------|-------------|------|
| opening a file should create a session | CreateEmpty → Open → verify sessionId | Integration + Excel |
| listing sessions should return active sessions | Verifies session fields: sessionId, filePath, canClose, isExcelVisible | Integration + Excel |
| closing a session should remove it from the list | Before/after session count comparison | Integration + Excel |
| session info should include all required fields for QuickPick | Validates all fields used by showSessionsQuickPick | Integration + Excel |

### Unit Tests (`tests/unit/`)

| Test File | Coverage |
|-----------|----------|
| mcpClient.test.ts | McpClient class, polling, error handling |
| statusBarMcp.test.ts | StatusBarMcp class, state management, visibility |
| showSessionsQuickPick.test.ts | QuickPick formatting, action handling |
| extension.test.ts | Extension activation, disposal |

### Test Totals
- **Integration Tests**: 16 tests (7 status bar + 5 server + 4 session lifecycle)
- **Unit Tests**: 29 tests
- **Total**: 45 tests

