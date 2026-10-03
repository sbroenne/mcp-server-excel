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
rows. Append uses existing column order. Every row in `rows` or `rows_file`
(MCP) / `--rows` or `--rows-file` (CLI) must have exactly the Table's column
count; include `null` for an intentionally blank cell. All row widths are checked
before writing cells or changing calculation mode. A mismatched row is rejected
without truncating values, inserting incomplete records, or expanding the Table.
This validation does not promise rollback for unrelated later Excel failures.

## Native filtering

MCP `table_column` / CLI `tablecolumn` `apply-filter` uses typed `options` /
`--options`. The former `apply-filter-values` action is removed, not an alias.
Other column filters are preserved. `clear-filters` affects the selected Table,
not every worksheet filter.

```mcp
table_column(action: 'apply-filter', session_id: sessionId, table_name: 'Sales', column_name: 'Amount', options: {filterOperator: 'And', criteria1: '>=100', criteria2: '<=500'})
table_column(action: 'apply-filter', session_id: sessionId, table_name: 'Sales', column_name: 'Region', options: {filterOperator: 'Values', values: ['North','West','Central']})
```

```cli
excelcli -q tablecolumn apply-filter --session $sessionId --table-name Sales --column-name Amount --options '{"filterOperator":"And","criteria1":">=100","criteria2":"<=500"}'
excelcli -q tablecolumn apply-filter --session $sessionId --table-name Sales --column-name Region --options '{"filterOperator":"Values","values":["North","West","Central"]}'
```

Nested option keys stay camelCase. Native date groups select a year/month/day,
not a text-only date comparison. Color, icon, top/bottom, and dynamic operators
require only their applicable typed fields; unrelated settings fail before writes.
`get-filters` returns every column, native operators, and both criterion slots.
Arrays stay arrays. Excel can normalize two selected values into an OR condition;
reads report native state, not a reconstruction of the original request.
Criterion getter failures include their HRESULT and `readError`, not invented
blank criteria. A native empty variant has `emptyVariant: true`.

Use ordinary-range filtering and criteria-range filtering through
[range operations](range.md), not Table creation as a prerequisite.

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

`update-dax` stores the new connection command before executing it. If the
query fails, do not assume the original command was restored: `get-dax` reports
the stored command, while worksheet rows can still contain the last successful
result. Correct the query and call `update-dax` again to establish a successful
refresh.
