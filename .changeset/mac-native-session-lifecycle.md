---
"excelmcp": patch
---

Use direct native Apple Events for experimental macOS workbook opening and
closing, preserving exact-workbook ownership, save/discard behavior, and
explicit timeout recovery through both the MCP Server and `excelcli`.

Use shared generated command routing for both platforms. Add build and
compatibility-check infrastructure for an independently versioned, optional
macOS helper add-in; helper-backed public actions remain unavailable until
desktop acceptance is complete.

Create macOS worksheets through native events, inserting before the active
sheet to match Windows instead of appending at the end.

Match Windows exact-name lookup for worksheet rename/delete, rejecting missing,
case-mismatched, or additionally quoted names before mutation with the shared
Core diagnostic.

Validate new worksheet names through shared Windows/Mac rules before mutation,
including reserved names and case-insensitive duplicates, without trimming.
Workbooks containing chart sheets remain an explicitly gated naming variant.

Rename worksheets through native object references. Preserve shared Excel
naming-rejection diagnostics and the original native error through the child
boundary, without converting transport failures or timeouts into naming errors.

Read and write rectangular A1/R1C1 formulas through native Apple Events with
shared validation and error diagnostics. Preserve blank-valued formula
occupancy, default overwrite rejection, and save/reopen behavior. Normalize
Excel's horizontal evaluation vectors without accepting malformed range shapes.
Preserve JSON-file A1/R1C1 inputs, dynamic spills, and explicit `@` formulas
through both entry points and save/reopen.
Discover formula errors before reading non-error runs, avoiding the native
bulk Value2 getter that terminated Excel for mixed error matrices.

Calculate normal sheet and rectangular range scopes through native Apple
Events without changing calculation mode or active worksheet, or calculating
dirty formulas in unrelated scopes and workbooks. Application scope,
full/rebuild, and unsupported address variants remain explicitly gated.
Preserve explicit unsupported-variant error categories from native automation
children instead of misreporting them as COM failures.
