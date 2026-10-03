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

Choose the output according to the task. Saving under a new name changes the
active workbook's path; saving a same-format copy leaves the active workbook
unchanged. PDF/XPS output is a published view, not an editable workbook.

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
