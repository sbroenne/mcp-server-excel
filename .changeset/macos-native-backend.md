---
"excelmcp": minor
---

Add the first capability-gated macOS release with a native Excel Apple Events
backend while preserving the existing Windows COM backend. The initial
slice supports session lifecycle, worksheet listing/rename/delete, core range
values and formulas, clears, and calculation through both MCP and `excelcli`.
Unsupported macOS operations fail explicitly.

Existing-file workflows use LaunchServices after a non-prompting Automation
permission check and reject workbook-name collisions before Excel can display a
modal dialog. Mac daemon output is isolated from CLI command output, and MCP
path errors use platform-appropriate wording. macOS ARM64 CLI and MCP archives
are included in releases. Power Query and VBA remain capability-gated while the
validated Power Query package path and unresolved VBA recompilation/trust path
are implemented.
