# Workbook Lifecycle

Use the `workbook` tool or CLI command group for workbook-level metadata, integrity checks, file variants, publishing, and external links. Use `file` only for opening, creating, listing, and closing sessions.

## Metadata and document properties

- `get-info` returns the active workbook name, path, Excel file format, saved/read-only state, and password/write-reservation flags.
- `list-document-properties` can include built-in properties, custom properties, or both.
- `get-document-property` and `set-document-property` require `scope`: `built-in` or `custom`.
- Built-in properties can be read and updated but not deleted.
- Missing custom properties are created as string properties by `set-document-property`; `delete-document-property` removes custom properties only.

## Integrity validation

- `validate-integrity` is read-only. It does not calculate, refresh, update links, edit, or save the workbook.
- Omit `checks` to inspect formula errors, external links, and tables; caller-supplied control totals are also checked when present. Use `worksheet_names` to limit formula and table scans in large workbooks.
- Control-total entries require `sheetName`, single-cell `cellAddress`, finite numeric `expectedValue`, and an optional non-negative absolute `tolerance`.
- Run `calculation_mode calculate` first when current calculated results are required. Validation reports manual or incomplete calculation state instead of changing it.
- `success` means the validation operation completed; use `overallStatus` for workbook integrity. `failed` means at least one error; `passed-with-warnings` contains only warnings; `passed` has neither. Findings are grouped by severity/category. Formula, link, table-structure, header, and control-total findings are deterministic for current workbook state; calculated-column findings are heuristic.
- `max_findings` limits returned details, not counts. Check `findingsTruncated` before assuming every finding is present in the response.

## Save and publish

- `save-as` supports `auto`, `xlsx`, `xlsm`, `xlsb`, and `xls`. The file extension must match the selected format, and the active session follows the new path.
- `save-copy-as` preserves the current format and leaves the active workbook/session unchanged. Its target extension must match the active workbook.
- `export-fixed-format` publishes PDF or XPS. Keep `open_after_publish=false` for unattended workflows.
- Output directories must already exist. Existing files require `overwrite=true`.

Changing formats can remove unsupported workbook features. In particular, saving a macro-enabled workbook as `.xlsx` removes VBA content after Excel's format conversion.

## External Excel links

1. Call `list-external-links` and use the exact returned `source`.
2. Call `update-external-link` to refresh one source.
3. Call `break-external-link` only with explicit user intent: it permanently replaces linked formulas with their current values.

Printing and print preview are not exposed. Printing can send output to a physical default printer, and preview is modal and can block unattended Excel sessions.
