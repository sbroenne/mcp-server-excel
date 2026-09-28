---
"excelmcp": minor
---

Add the first capability-gated macOS release with a native Excel Apple Events
backend while preserving the existing Windows COM backend. The initial
slice supports session lifecycle, worksheet creation/listing/rename/deletion,
core range values, formulas, number formats, explicit row/column sizing, clears,
and calculation through both MCP and `excelcli`. Unsupported macOS operations
fail explicitly.

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

Add native macOS range parity for complete, values-only, and formulas-only
copy; two-dimensional number formats; range information; row and column
auto-fit; merge and unmerge; and cell lock get/set through both MCP and
`excelcli`. Used-range, current-region, and merged-area discovery remain
explicitly unavailable because Excel's declared Apple Events routes did not
preserve their Windows contracts.

Return explicit recovery guidance when a macOS file-open handoff cannot be
confirmed. Preserve newly created workbooks instead of deleting them while
Excel may still open them. Reject queued operations and avoid automatic close
on uncertain sessions.

Preserve complete helper Power Query view and load-configuration results in the
shared Mac response serializer, including M code, query identity and load
destination, instead of returning only base success and file-path fields.

Fix daemon startup when invoking `dotnet excelcli.dll`: the child process now
receives the CLI assembly path instead of attempting to execute `dotnet service`.
Use a macOS-compatible private pipe name and bounded native CLI shutdown in the
Power Query acceptance runner rather than Windows-only cleanup discovery.

Treat workbook files as opaque and route Power Query only through the optional
trusted helper.
