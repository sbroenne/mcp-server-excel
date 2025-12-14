# Data Model: Status Bar MCP Server Monitor

## Entities

- MCP Server State
  - Fields: `isRunning` (bool), `lastChecked` (datetime), `connectionStatus` (enum: running, stopped, unreachable)

- MCP Session
  - Fields: `sessionId` (string), `filePath` (string), `fileName` (string), `createdAt` (datetime), `lastActivityAt` (datetime), `hasUnsavedChanges` (bool)

## Relationships

- MCP Server State 1 — N MCP Session (server hosts multiple sessions)

## Validation Rules

- `fileName` is derived from `filePath` (basename)
- `connectionStatus` maps from MCP Server tool response (success flag + reachability)
- `hasUnsavedChanges` if exposed by server; otherwise omitted

## State Transitions

- Server: stopped → running (on successful poll), running → stopped (on failure), running → unreachable (on timeout)
- Session: active → closing (user action), closing → closed (success), closing → error (failure)
