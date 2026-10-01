# Worksheet Tables versus model tables

Create a worksheet Table when requested or required, not for every rectangular
dataset. Reuse existing Tables: append, resize, or update rather than recreate.

## Conversion and preservation

Use `preflight` when conversion boundaries or sorting safety are uncertain.
Create enforces the same blockers, but warnings remain advisory. A
`safeToCreate` result does not establish every business/layout consequence.
For headerless data use `has_headers: false` (MCP) / `--has-headers false` (CLI).

`delete` converts a Table to a plain range and keeps cell data; it can still
break dependent PivotTables/model objects. Shrinking changes membership, not
permission to clear excluded cells.

`read` is metadata; `get-data` is cell values. Ordinary reads include filtered
rows. Use `visible_only: true` (MCP) / `--visible-only true` (CLI) for visible
rows. Append uses existing column order.

## Styling

Use `table_style` (MCP) / `--table-style` (CLI) at creation or `set-style`
later. The Table owns its visual style, not `range_format` (MCP) /
`rangeformat` (CLI). Column number formats and totals functions are separate.
For a requested style change on existing `Sales`:

```mcp
table(action: 'set-style', session_id: sessionId, table_name: 'Sales', table_style: 'TableStyleMedium2')
```

```cli
excelcli -q table set-style --session $sessionId --table-name Sales --table-style TableStyleMedium2
```

## Model and worksheet results

A worksheet Table is **not automatically in Power Pivot**.
`add-to-data-model` adds an existing Table and is idempotent. Power Query can
load directly with `load_destination: 'data-model'` (MCP) /
`--load-destination data-model` (CLI).
See [model prerequisites and refresh](datamodel.md).

`create-from-dax` creates a worksheet Table from a model `EVALUATE` query.
`get-dax` inspects it; `update-dax` changes it. This is a worksheet result,
not DAX calculated-table creation in the model. Use model `evaluate` to return
results without a worksheet object, or a [PivotTable](pivottable.md) for
interactive filtering.
