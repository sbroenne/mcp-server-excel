# Worksheets

Same-workbook lifecycle operations create, list, rename, copy, move, and delete
sheets. Always use the captured session. Rename requires the old and new names,
not the source/target parameters used for copying.

```mcp
worksheet(action: 'rename', session_id: sessionId, old_name: 'Sheet1', new_name: 'Summary')
```

```cli
excelcli -q sheet rename --session $sessionId --old-name Sheet1 --new-name Summary
```

The server checks Microsoft's documented naming rules before creating, copying,
or renaming a worksheet: names cannot be blank, exceed 31 characters, contain
`/ \ ? * : [ ]`, begin or end with an apostrophe, or be `History`.
Non-English names and apostrophes inside a name are allowed. See Microsoft's
[worksheet naming guidance](https://support.microsoft.com/en-us/excel/rename-a-worksheet).
Names are not trimmed: leading or trailing spaces may be part of a name, but a
name made entirely of whitespace is rejected as blank.
The server checks these documented constraints before changing the workbook;
Excel also validates the name when it is assigned. If Excel rejects a name
after a create or copy has added the sheet, the error identifies that sheet.
It remains in the workbook for the caller to inspect and remove if appropriate.

For ordering, specify before **or** after another sheet, not both. Inspect names
and dependencies before deleting or replacing anything.
Deleting a sheet removes all of its contents and can break dependent references.
There is no tool-level undo. Discarding unsaved changes also discards earlier
unsaved work.

## Cross-file operations

Copy-to-file and move-to-file manage opening, saving, and closing files in one
call without a supplied session. Use the exact user-provided source and target
paths. Do not apply them to files already open in unrelated sessions.

```mcp
worksheet(action: 'copy-to-file', source_file: sourcePath, source_sheet: 'Summary', target_file: targetPath, target_sheet_name: 'Q1 Summary')
```

```cli
excelcli -q sheet copy-to-file --source-file $sourcePath --source-sheet Summary --target-file $targetPath --target-sheet-name 'Q1 Summary'
```

Cross-file copy can rename the copied sheet. Both transfer operations support
positioning relative to a target sheet. Same-file copying uses the ordinary copy
action instead. Do not assume a failure rolls back every file; inspect both
files before retrying a transfer.
Move-to-file removes the source sheet and saves both workbooks. There is no
tool-level undo. Closing another session without saving cannot reverse that
saved transfer.

## Styling and outlines

Worksheet-style operations own tab colors, visibility, protection, page setup,
legacy notes, and row/column grouping. Legacy notes are not threaded comments;
use range-link operations for threaded comments.

Page setup never prints or opens a preview. Use `worksheet_style` actions
`set-page-setup` / `get-page-setup` (CLI `worksheetstyle set-page-setup` /
`worksheetstyle get-page-setup`). Additional settings use MCP `page_setup_options` or
CLI `--page-setup-options`; nested keys remain camelCase. Margins are **points**,
not inches. Omitted/null settings remain unchanged; empty text explicitly clears
`printArea`, `printTitleRows`, `printTitleColumns`, or header/footer text.
Print titles select complete rows/columns. Printer support is required for
paper settings and automatic pagination.

`zoomPercent` selects fixed scaling and cannot accompany MCP
`fit_to_pages_wide` / `fit_to_pages_tall` (CLI `--fit-to-pages-wide` /
`--fit-to-pages-tall`). Fit counts of zero mean unlimited in that dimension.
Reads report null `zoomPercent` in fit mode and null fit counts when fixed zoom
is active; an unlimited fit dimension is also null.

`set-page-breaks` (CLI `worksheetstyle set-page-breaks`) takes MCP
`page_break_options` or CLI `--page-break-options`,
with required `rows` and `columns` lists of positions **before which** to break.
It replaces **all manual breaks on the worksheet**, including outside the
current print area; empty lists clear them. `get-page-breaks` reads every break
Excel exposes in the current print scope, identifies manual/automatic breaks,
and states that scope limitation. Scaling can override manual breaks in output.
Use `workbook export-fixed-format` / MCP `workbook` `export-fixed-format` for
PDF/XPS output that honors these settings by default. Native failures do not
promise rollback of settings already changed.

Group complete rows such as `2:10` with axis Rows, or columns such as `B:F` with
axis Columns. Summary rows use above/below and summary columns left/right.
Read outline information before changing it, use show-outline-levels to
expand/collapse, ungroup to remove a level, and clear-outline to remove all groups.

## Protection permissions

`worksheet_style` (MCP) / `worksheetstyle` (CLI) `set-protection` accepts an
`options` object with native permissions. Supplied options replace the
configuration; unspecified permissions use Excel's restrictive defaults,
not the previous permissions. Options are not valid when unprotecting.
`get-protection` reads protected components, every permission flag, and
selection restrictions without returning passwords.

```mcp
worksheet_style(action: 'set-protection', session_id: sessionId, sheet_name: 'Summary', is_protected: true, options: {"allowFiltering": true, "allowFormattingRows": true, "userInterfaceOnly": true})
worksheet_style(action: 'get-protection', session_id: sessionId, sheet_name: 'Summary')
```

```cli
excelcli -q worksheetstyle set-protection --session $sessionId --sheet Summary --is-protected true --options '{"allowFiltering":true,"allowFormattingRows":true,"userInterfaceOnly":true}'
excelcli -q worksheetstyle get-protection --session $sessionId --sheet Summary
```

Nested option keys keep the spelling shown above in both entry points. Filtering permission
allows changing an existing filter, not creating/removing it. Sorting/deleting
still requires unlocked cells. UI-only protection permits automation but is
runtime-only: it is not retained after reopening. Reapply it only when wanted.
Sheet protection and formula hiding are not file encryption or a confidentiality
boundary; native failures do not promise rollback.

## Workbook themes

`workbook` `get-theme` reads all native theme color slots and major/minor
Latin, East Asian, and complex-script font definitions. Empty font names are
native unspecified definitions, not fallback fonts. Excel COM does not expose
the applied theme file/name or scheme names; the result states this limitation.
`apply-theme` changes theme-sensitive formatting throughout the workbook,
including theme-bound cell/style/chart colors and fonts. Explicit RGB colors
remain fixed. Use an existing theme file selected for the requested appearance;
do not create workbook copies as a prerequisite. Saving remains explicit.

```mcp
workbook(action: 'get-theme', session_id: sessionId)
workbook(action: 'apply-theme', session_id: sessionId, theme_path: themePath)
```

```cli
excelcli -q workbook get-theme --session $sessionId
excelcli -q workbook apply-theme --session $sessionId --theme-path $themePath
```

The input path must be absolute and select an existing Office `.thmx` file.
Native failures do not promise rollback; read the actual theme definitions
before retrying.
