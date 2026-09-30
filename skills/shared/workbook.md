# Workbook Lifecycle

Use workbook operations for metadata, file variants, publishing, and external
links. Session lifecycle is separate.

```mcp
workbook(action: 'get-info', session_id: sessionId)
workbook(action: 'list-external-links', session_id: sessionId)
```

```cli
excelcli -q workbook get-info --session $sessionId
excelcli -q workbook list-external-links --session $sessionId
```

## Metadata and document properties

- `get-info` returns the active workbook name, path, Excel file format, saved/read-only state, and password/write-reservation flags.
- `list-document-properties` can include built-in properties, custom properties, or both.
- `get-document-property` and `set-document-property` default to custom properties.
  Select `built-in` scope explicitly when accessing a built-in property.
- Built-in properties can be read and updated but not deleted.
- Missing custom properties are created as string properties by `set-document-property`; `delete-document-property` removes custom properties only.

## Save and publish

- `save-as` supports `auto`, `xlsx`, `xlsm`, `xlsb`, and `xls`. The file extension must match the selected format, and the active session follows the new path.
- `save-copy-as` preserves the current format and leaves the active workbook/session unchanged. Its target extension must match the active workbook.
- `export-fixed-format` publishes PDF or XPS. Leave the open-after-publish option
  disabled (the default) for unattended workflows.
- Output directories must already exist. Replacing an existing output file
  requires the overwrite option and authorization for that replacement.

Changing formats can remove unsupported workbook features. In particular, saving a macro-enabled workbook as `.xlsx` removes VBA content after Excel's format conversion.

## External Excel links

1. Call `list-external-links` and use the exact returned `source`.
2. Call `update-external-link` to refresh one source.
3. Call `break-external-link` only with explicit user intent: it permanently replaces linked formulas with their current values.

Printing and print preview are not exposed. Printing can send output to a physical default printer, and preview is modal and can block unattended Excel sessions.
