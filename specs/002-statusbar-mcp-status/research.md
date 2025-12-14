# Research: Status Bar MCP Server Monitor

Date: 2025-12-14  
Branch: 002-statusbar-mcp-status

## Decisions

- Interaction model: Click opens Quick Pick; hover is read-only tooltip
- Status polling: Poll MCP Server status every 3s with backoff to 5s on error; update immediately on successful actions
- Session listing: Query MCP Server sessions endpoint; present file name (basename) as label, full path as detail, inactivity indicator via last activity timestamp when available
- Actions: Implement Close and Save & Close via MCP server tools; return JSON, show result in notification

## Rationale

- VS Code tooltips are read-only; Quick Pick provides accessible action UI without complex webview
- Polling aligns with spec’s 5s update requirement and avoids heavy event subscription complexity
- Basename improves readability; full path remains available to disambiguate
- MCP Server’s JSON contract ensures reliable error handling and messaging

## Alternatives Considered

- Webview Panel for sessions: richer UI, but heavier and overkill for simple actions
- Status bar cycling on click: compact but confusing; lower discoverability of per-session actions
- Event-driven updates via MCP notifications: not currently supported; would require protocol extension

## Open Questions (resolved)

- Automated UI testing: keep manual for now; add smoke tests later if extension test harness is available
- Server endpoints: reuse existing Excel MCP tools to list sessions via `excel_file(action: 'list')`; add minimal helper in extension to format results
