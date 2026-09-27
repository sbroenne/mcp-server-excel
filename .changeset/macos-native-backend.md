---
"mcp-server-excel": minor
---

Add an experimental, capability-gated macOS host and native Excel Apple Events
backend while preserving the existing Windows COM backend. The first vertical
slice supports session lifecycle, worksheet listing/rename/delete, core range
values and formulas, calculation, and workbook-qualified VBA procedure execution
through both MCP and `excelcli`. Unsupported macOS operations fail explicitly.

Existing-file workflows use LaunchServices after a non-prompting Automation
permission check and reject workbook-name collisions before Excel can display a
modal dialog. Mac daemon output is isolated from CLI command output, and MCP
path errors use platform-appropriate wording. Power Query and VBA source
editing remain production-gated while the validated Power Query package path
and unresolved VBA project recompilation path are implemented.
