# Workbook Lifecycle

Inspect the intended workbook's saved/read-only state, properties, and external
dependencies before changing its format or publishing output. Session lifecycle
is separate; see [session and saving guidance](behavioral-rules.md#sessions-and-failures).

Current commands and supported inputs come from CLI help or MCP tool descriptions.

## Opening a SharePoint workbook

Open accepts a direct SharePoint or OneDrive for Business HTTPS file URL ending
in `.xlsx`, `.xlsm`, `.xlsb`, or `.xls`. An optional `?web=1` or `?web=0` is
removed before opening and session matching. Spaces and their percent-encoded
forms identify the same workbook. Folder URLs, browser pages such as `Doc.aspx`,
sharing links, other query parameters, arbitrary websites, and OneDrive personal
consumer links are not supported.

Use MCP `file(action: 'open', file_path: '<direct-url>', show: true)` or
`excelcli session open "<direct-url>" --show`. The visible session is required
so Office sign-in and rights-management prompts remain accessible. Excel uses
the signed-in Office account; ExcelMcp neither supplies credentials nor bypasses
IRM/AIP restrictions. Inspect `workbook_read` action `get-info` or
`excelcli workbook get-info` for `readOnly` and the live `autoSaveOn` status before
editing. When `autoSaveOn` is true, Excel can persist changes without an explicit
save.

Excel versions without the AutoSave property report `autoSaveOn: false`; there
is no AutoSave feature to disable on those versions. Only known unavailable-member
errors are treated this way. Other AutoSave read or write failures remain errors.

`file_read` action `test` and `excelcli session test` report that URL validation
requires interactive opening. They do not download the workbook or probe a local
file. Returned `exists: false`, zero size, and `isIrmProtected: false` are unknown
remote metadata, not evidence of a missing or unprotected workbook.

Cloud AutoSave is disabled in editable remote sessions. Explicit save/close
behavior therefore matches local workbooks: close without saving discards
unsaved session edits; close with saving writes through Excel to the original
SharePoint location. Normal Service shutdown still attempts to save. A successful
save confirms Excel's saved state, not independent server-side upload completion.
Save As, Save Copy As, export, and create continue to require Windows output paths.

## Metadata and document properties

Use built-in properties for existing document metadata and custom properties
for workbook-specific information. Custom properties can be removed; built-in
ones are maintained by Excel. Avoid putting private paths, credentials, or
connection strings in published properties.

Use MCP `workbook_read` action `inspect` or `excelcli workbook inspect` for a
bounded workbook overview. It returns worksheet visibility and used-range
geometry, table locations, and visible user-defined names, with per-section
counts and omitted-item counts. `sheet_name` narrows worksheet and table
metadata; the named-range section remains workbook-wide. Hidden worksheets are
reported, not skipped. The operation reads through Excel and does not change the
workbook or its view.

Cell previews are optional. Set `include_preview=true` and provide `sheet_name`;
optionally set `range_address` to limit the source range. Without it, the
preview starts at the top-left of Excel's UsedRange. The preview reads at most
10 rows by 10 columns, and reports omitted rows and columns. `max_cell_characters`
and `max_preview_characters` further bound text in returned values and formulas.
Row and column limits are applied before reading cell contents; text limits are
applied to the response. For example:

```text
MCP: workbook_read(action: 'inspect', workbook_session_id: sessionId,
     sheet_name: 'Summary', include_preview: true, range_address: 'A1:F100',
     max_preview_rows: 5, max_preview_columns: 6)
CLI: excelcli -q workbook inspect --session <session-id> --sheet-name Summary
     --include-preview true --range-address A1:F100 --max-preview-rows 5
     --max-preview-columns 6
```

Excel reports an empty worksheet's UsedRange as `$A$1`. Its preview contains a
null value and an empty formula; the returned address does not mean the cell is
populated.

If Excel cannot read a protected or otherwise inaccessible range, the operation
reports an error rather than presenting it as empty.

## Save and publish

Opening or creating a workbook requires a confirmed Excel process ID and start
time for safe readiness checks and cleanup. Startup retries a temporary capture
failure. If the identity remains unavailable, it reports an error before opening
or creating the workbook rather than returning a session that cannot save or close.

Inspect `readOnly` with MCP `workbook_read` action `get-info` or
`excelcli workbook get-info` before editing. Protected workbooks require visible
authentication; Excel decides editing rights. Do not change protection to work
around a genuine permission restriction.

Workbook-changing actions reject read-only access before editing. Inspection,
window controls, calculation other than workbook precision changes, and authorized
Save As/copy/export remain available. MCP `calculation_mode` action `set-precision`
and `excelcli calculation set-precision` require editing rights.
Excel still enforces permissions on outputs.

Writes change the open workbook, not necessarily the saved file. Saving a
read-only workbook or a save cancelled by Excel returns an error. A failed save
leaves the session open and any unsaved changes available for inspection; do not
assume they were persisted or automatically discard them.

Successful saving confirms Excel's saved state, not completion of OneDrive
synchronization or upload to SharePoint.

Saving, Save As, and explicit closing check Excel's live refresh readiness.
A running refresh, open modal dialog, busy Excel instance, or unconfirmed status returns a `Busy`
error without saving or discarding edits. The session remains open. Wait for
Excel to finish, check `canClose` through MCP `file_read` action `list` or
`excelcli session list`, inspect the refreshed values, then retry. A zero
`activeOperations` count alone is not proof that Excel is idle.

The same session listing exposes `excelState` and `blockingReason`. When
`excelState` is `dialogOpen`, check the Excel window for a prompt before simply
waiting longer. The server observes window ownership without reading dialog
contents; this can detect a separate Microsoft sign-in host but cannot confirm
that authentication is the dialog's purpose. It never enters credentials or
responds to the prompt. After responding, check readiness and refreshed data
before saving.

Power BI/MSOLAP (OLAP) connections always refresh synchronously. Their unsupported
background setting is reported as false; enabling it is rejected before changing
other properties. An empty last-refresh date is reported as unknown, not as proof
of a completed query.

Choose the output according to the task. Saving under a new name changes the
active workbook's path; saving a same-format copy leaves the active workbook
unchanged. PDF/XPS output is a published view, not an editable workbook.

Saved `.xlsx`, `.xlsm`, `.xlsb`, and `.xls` workbooks can be reopened with
`file(action: 'open', file_path: ...)` (MCP) or `excelcli session open <FILE>` (CLI).

Replacing an output needs authorization. A format conversion can remove
features: saving a macro-enabled workbook as `.xlsx` removes its VBA content.
Inspect the result and active path before continuing work.

For unattended exports, avoid opening viewers or modal previews. Configure
the [print layout](worksheet.md#styling-and-outlines) before publishing.

## External Excel links

Discover the actual linked sources before refreshing them. A refresh can change
calculated results; it is not merely inspecting the workbook.

Breaking a link permanently replaces linked formulas with their current values.
Do it only when that change is intended, and inspect the resulting values before
saving. There is no tool-level undo. Closing without saving loses all earlier
unsaved work too.

Printing and print preview are not exposed: a default printer can produce
physical output, and modal preview can block unattended sessions.
