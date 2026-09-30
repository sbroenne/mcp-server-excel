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

For ordering, specify before **or** after another sheet, not both. Inspect names
and dependencies before deleting or replacing anything.

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

## Styling and outlines

Worksheet-style operations own tab colors, visibility, protection, page setup,
legacy notes, and row/column grouping. Legacy notes are not threaded comments;
use range-link operations for threaded comments.

Group complete rows such as `2:10` with axis Rows, or columns such as `B:F` with
axis Columns. Summary rows use above/below and summary columns left/right.
Read outline information before changing it, use show-outline-levels to
expand/collapse, ungroup to remove a level, and clear-outline to remove all groups.
