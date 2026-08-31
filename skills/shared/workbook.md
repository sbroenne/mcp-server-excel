# Workbook Lifecycle

Use the `workbook` tool or CLI command group for workbook-level metadata, file variants, publishing, and external links. Use `file` for session lifecycle and session-owned savepoints.

## Metadata and document properties

- `get-info` returns the active workbook name, path, Excel file format, saved/read-only state, and password/write-reservation flags.
- `list-document-properties` can include built-in properties, custom properties, or both.
- `get-document-property` and `set-document-property` require `scope`: `built-in` or `custom`.
- Built-in properties can be read and updated but not deleted.
- Missing custom properties are created as string properties by `set-document-property`; `delete-document-property` removes custom properties only.

## Save and publish

- `save-as` supports `auto`, `xlsx`, `xlsm`, `xlsb`, and `xls`. The file extension must match the selected format, and the active session follows the new path.
- `save-copy-as` preserves the current format and leaves the active workbook/session unchanged. Its target extension must match the active workbook.
- `export-fixed-format` publishes PDF or XPS. Keep `open_after_publish=false` for unattended workflows.
- Output directories must already exist. Existing files require `overwrite=true`.

Changing formats can remove unsupported workbook features. In particular, saving a macro-enabled workbook as `.xlsx` removes VBA content after Excel's format conversion.

## Savepoints and rollback

- `create-savepoint` captures the current unsaved serializable workbook state without changing the workbook path or saved flag.
- `rollback-savepoint` persistently restores that snapshot to the same path while keeping the public session ID. The savepoint remains available until `release-savepoint`.
- `list-savepoints` reports retained snapshot sizes and limits. Each session can retain 8 savepoints and 1 GiB; one service process can retain 4 GiB.
- A successful `save-as` releases every savepoint for that session. Savepoints never move a session back to an earlier path.
- Savepoints reject read-only/IRM workbooks, active calculation or refresh, refresh-on-open connections, and connection types whose refresh state ExcelMcp cannot verify safely.

Savepoints cover workbook state that Excel serializes, including supported VBA, Power Query, Data Model, connection, table, chart, PivotTable, formula, and formatting state. They do not undo effects outside the workbook, such as database writes, exported files, network calls, printing, or VBA changes to other systems. Volatile formulas and external data can recalculate after rollback.

The limits apply to retained savepoints. Rollback also needs temporary space for an emergency copy of the current workbook plus a same-volume replacement copy; ExcelMcp checks free space before closing the live workbook and removes those files after recovery stabilizes.

## External Excel links

1. Call `list-external-links` and use the exact returned `source`.
2. Call `update-external-link` to refresh one source.
3. Call `break-external-link` only with explicit user intent: it permanently replaces linked formulas with their current values.

Printing and print preview are not exposed. Printing can send output to a physical default printer, and preview is modal and can block unattended Excel sessions.
