---
"excelmcp": minor
---

Add the first capability-gated macOS release with a native Excel Apple Events
backend while preserving the existing Windows COM backend. The initial
slice supports session lifecycle, worksheet creation/listing/rename/deletion,
core range values, formulas, number formats, explicit row/column sizing, clears,
and calculation through both MCP and `excelcli`. Saved, clean workbooks also
support Power Query list/view with exact M and worksheet load-state inspection.
Unsupported macOS operations fail explicitly.

Existing-file workflows use LaunchServices after a non-prompting Automation
permission check and reject workbook-name collisions before Excel can display a
modal dialog. Mac daemon output is isolated from CLI command output, and MCP
path errors use platform-appropriate wording. macOS ARM64 CLI and MCP archives
are included in releases, along with separate native Windows and Apple Silicon
macOS Claude Desktop MCPB bundles. Power Query mutations and VBA remain
capability-gated while transactional workbook orchestration and the optional
macro/VBA trust tiers are implemented.

Correct the guarded screenshot process lookup to load AppKit and normalize
Objective-C collection counts before selecting the exact Excel process.

Add guarded native named-range lifecycle and bounded visible-name preview
implementations, with an explicit CLI/MCP acceptance workflow. These candidate
commands remain unavailable by default until desktop behavior is verified.
