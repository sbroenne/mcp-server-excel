---
name: excel-cli
description: >
  Quick start for changing or checking an Excel workbook with the excelcli
  command. Use before your first excelcli call. Gives the usual open, change,
  check and save routine and the exact batch names for common actions, so you
  can skip most --help and list-actions lookups.
---

# excelcli quick start

If `excelcli` is not on PATH, run this plugin's `bin\start-cli.ps1` by its full
path with the same arguments.

## The routine: about three calls

1. **Open and look** in one batch. Start the batch with `session.open`; later
   entries in the same batch use that workbook automatically. Read what you
   need to plan, such as `sheet.list`, `table.list` and `range.get-used-range`.
   The result of `session.open` gives the `sessionId`. Before editing, check
   its `autoSaveOn` and `warning`: Excel may save edits automatically.
2. **Change and check** in one batch with `--session <sessionId>
   --stop-on-error`. Put the writes first, then reads of the cells you changed,
   so one result shows that the change worked.
3. **Save and close**: `session close --session <sessionId> --save`, only after
   step 2's reads look right. Closing without `--save` discards only unsaved
   edits; it cannot undo changes already saved by AutoSave.

Pass the batch as a JSON array in a file with `--input <file>`. Direct
`excelcli` calls also accept standard input (`batch` with no `--input`), but
the PowerShell launcher requires `--input <file>`. Each entry is
`{"command":"group.action","args":{...}}`. Argument names are case-sensitive
camelCase. Put arrays and objects in as JSON, not as strings of JSON. Check
every entry's `success` in the output; one failed entry does not undo the
others.

```json
[
  {"command":"session.open","args":{"filePath":"C:\\work\\prices.xlsx"}},
  {"command":"sheet.list"},
  {"command":"table.list"},
  {"command":"range.get-used-range","args":{"sheetName":"Prices"}}
]
```

Keep reads small: ask for the rows and columns you need, not a whole large
sheet. Don't close and reopen the workbook to check your work; read it back in
the same session.

## Exact names for common actions

Only run `<group> <action> --help` for actions that are not listed here.

| Command | args |
|---|---|
| `session.open` | `filePath` |
| `session.close` | `save` (true/false) |
| `sheet.list` | none |
| `sheet.create` | `sheetName` |
| `range.get-used-range` | `sheetName` |
| `range.get-values` | `sheetName`, `rangeAddress` |
| `range.get-formulas` | `sheetName`, `rangeAddress` |
| `range.get-info` | `sheetName`, `rangeAddress` |
| `range.set-values` | `sheetName`, `rangeAddress`, `values` (2D array), `overwritePolicy` |
| `range.set-formulas` | `sheetName`, `rangeAddress`, `formulas` (2D array, each formula starts with `=`), `overwritePolicy` |
| `range.set-number-format` | `sheetName`, `rangeAddress`, `formatCode` |
| `rangeedit.fill` | `sheetName`, `rangeAddress`, `direction`, `overwritePolicy` |
| `rangeedit.sort` | `sheetName`, `rangeAddress`, `sortColumns` (`[{"columnIndex":1,"ascending":true}]`, 1-based within the range), `hasHeaders` |
| `rangeformat.format` | `sheetName`, `rangeAddresses` (array), `formatOptions` (object: `bold`, `fontColor`, `fillColor`, `numberFormat`, `horizontalAlignment`, `wrapText`, `borders`, ...) |
| `rangeformat.auto-fit-columns` | `sheetName`, `rangeAddress` |
| `table.list` | none |
| `table.read` | `tableName` |
| `table.append` | `tableName`, `rows` (2D array) |
| `tablecolumn.add-column` | `tableName`, `columnName`, `position` (1-based, optional) |
| `tablecolumn.set-column-number-format` | `tableName`, `columnName`, `formatCode` |
| `pivottable.list` | none |
| `pivottable.read` | `pivotTableName` |
| `pivottable.create-from-table` | `tableName`, `destinationSheet`, `destinationCell`, `pivotTableName` |
| `pivottablefield.add-row-field` | `pivotTableName`, `fieldName` |
| `pivottablefield.add-value-field` | `pivotTableName`, `fieldName`, `aggregationFunction` (Sum, Count, Average, Max, Min), `customName` |
| `pivottablefield.set-field-format` | `pivotTableName`, `fieldName`, `numberFormat` |
| `slicer.create-table-slicer` | `tableName`, `columnName`, `slicerName`, `destinationSheet`, `position` |
| `slicer.set-table-slicer-selection` | `slicerName`, `selectedItems` (array) |
| `slicer.create-slicer` (PivotTable) | `pivotTableName`, `fieldName`, `slicerName`, `destinationSheet`, `position` |
| `slicer.set-slicer-selection` (PivotTable) | `slicerName`, `selectedItems` (array) |
| `chart.create-from-range` | `sheetName`, `sourceRangeAddress`, `chartType` (e.g. ColumnClustered, Line, Pie, BarClustered), `chartName`, `targetRange` |
| `chartconfig.set-title` | `chartName`, `title` |
| `powerquery.list` | none |
| `powerquery.view` | `queryName` |
| `powerquery.update` | `queryName`, `mCode`, `refresh` |
| `workbook.refresh-all` | `timeoutSeconds` (optional) |
| `window.freeze-panes` | `sheetName`, `frozenRows`, `frozenColumns` |

`range.set-values` and `range.set-formulas` refuse to overwrite cells that
already have content. Set `"overwritePolicy":"allow"` only when you mean to
replace them.
