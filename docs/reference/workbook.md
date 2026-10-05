# Workbook Lifecycle

Inspect the intended workbook's saved/read-only state, properties, and external
dependencies before changing its format or publishing output. Session lifecycle
is separate; see [session and saving guidance](behavioral-rules.md#sessions-and-failures).

Current commands and supported inputs come from CLI help or MCP tool descriptions.

## Metadata and document properties

Use built-in properties for existing document metadata and custom properties
for workbook-specific information. Custom properties can be removed; built-in
ones are maintained by Excel. Avoid putting private paths, credentials, or
connection strings in published properties.

## Save and publish

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

Choose the output according to the task. Saving under a new name changes the
active workbook's path; saving a same-format copy leaves the active workbook
unchanged. PDF/XPS output is a published view, not an editable workbook.

Saved `.xlsx`, `.xlsm`, `.xlsb`, and `.xls` workbooks can be reopened with
`file(action: 'open', path: ...)` (MCP) or `excelcli file open --path ...` (CLI).

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
