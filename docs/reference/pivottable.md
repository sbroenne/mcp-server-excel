# PivotTables

Reuse a suitable existing PivotTable. Creating one does not configure its
rows, columns, values, or filters; choose those fields, then refresh and read
the actual summary.

Current commands and inputs come from CLI help or MCP tool descriptions.

| Source | Calculation approach |
|--------|----------------------|
| Worksheet Table or range | Ordinary aggregations and regular calculated fields |
| Data Model | DAX measures and model relationships |

## Calculated fields are calculations on aggregates

For a regular PivotTable with a numeric `Sales` field, doubling Sales is safe.
This example shows the distinction between defining a calculation, placing it
in Values, and checking the result:

```mcp
pivottable_calc(action: 'create-calculated-field', workbook_session_id: sessionId, pivot_table_name: 'SalesPivot', field_name: 'DoubleSales', formula: '=Sales*2')
pivottable_field(action: 'add-value-field', workbook_session_id: sessionId, pivot_table_name: 'SalesPivot', field_name: 'DoubleSales', aggregation_function: 'Sum')
pivottable(action: 'refresh', workbook_session_id: sessionId, pivot_table_name: 'SalesPivot')
pivottable_calc_read(action: 'get-data', workbook_session_id: sessionId, pivot_table_name: 'SalesPivot')
```

```cli
excelcli -q pivottablecalc create-calculated-field --session $sessionId --pivot-table-name SalesPivot --field-name DoubleSales --formula '=Sales*2'
excelcli -q pivottablefield add-value-field --session $sessionId --pivot-table-name SalesPivot --field-name DoubleSales --aggregation-function Sum
excelcli -q pivottable refresh --session $sessionId --pivot-table-name SalesPivot
excelcli -q pivottablecalc get-data --session $sessionId --pivot-table-name SalesPivot
```

The session, PivotTable, and source field must already exist. Check each result
before continuing. Do not use `=Quantity*UnitPrice` as a regular PivotTable
calculated field for line-item revenue: rows `(2,10)` and `(3,20)` have total
revenue 80, not the product of summed inputs, 150.

Calculate each source row's revenue and sum that column, or use a model measure:

```dax
SUMX(Sales, Sales[Quantity] * Sales[UnitPrice])
```

The table must be in the model first; see [Data Model guidance](datamodel.md).

## Show Values As is separate from aggregation

Sum/Average and Show Values As are independent choices. Select the exact
displayed Values instance, especially when several instances share a source
field. Removing an additional calculation need not change aggregation.

Differences from a previous item depend on current sort/filter order. Parent
percentages depend on the relevant row/column hierarchy. Do not turn native
blank/error results into zero.

Excel does not expose every model-backed base-field/item setting and can ignore
additional calculations on model measures. Use a source measure when native
settings cannot establish the result. Read back settings and displayed data.

## Refresh, layout, and grouping

Refresh sources before their summaries. Worksheet-source edits may need a model
refresh, followed by PivotTable refresh; Power Query refresh updates its model
load but does not replace the dependent PivotTable refresh.

- `sort-field` orders the selected field's labels in ascending or descending
  order, including Data Model fields identified by their exact CubeField name.
- `set-field-name` updates the displayed value-field caption when the selected
  source field is in Values; it does not rename the source worksheet column.
  Use MCP `field_name` / `custom_name` or CLI `--field-name` / `--custom-name`.
- `set-field-format` accepts invariant Excel number formats. A bare dollar sign
  remains a literal dollar sign, not the machine's regional currency. Escaped or
  quoted literals and bracketed currency/locale codes remain intact. Excel may
  return a normalized format with an escaped dollar sign. Use MCP `number_format`
  or CLI `--number-format`.

Choose Compact for a nested view or Tabular/Outline when separate field columns
and repeated labels suit the output. Use the PivotTable's own styles and field
number formats, not plain-cell styling that refresh can overwrite.

Grouping and calculated fields described here belong to regular PivotTables.
For model-backed summaries, use source grouping columns and DAX measures.
Drill-through on a regular value cell creates a new sheet of underlying rows;
it is not a read-only audit.

Use a verified live [PivotChart](chart.md) when a chart must follow fields and
filters, rather than a regular chart of displayed cells.

## Native filters, layout, and source changes

Manual item visibility and calculated filters are different. Clearing calculated
filters need not clear manual selections or other fields. Use the actual placed
field and displayed Values captions rather than guessing from source names.
Multiple filters are not automatically authorized by adding another criterion.

Repeated labels require a noncompact layout. Expand/collapse targets one
visible parent item, not every member. These calculated-filter and item-expansion
features do not cover provider-dependent Data Model behavior.

Inspect the source, shared cache users, and connected controls before changing
data or cache settings. A supported regular source replacement retains the
same field schema and isolates the selected PivotTable; connected slicers and
timelines must first be deliberately disconnected. `set-source` does not
replace external or workbook Data Model sources.

Use `get-connection` to inspect a PivotTable's connection and `set-connection`
to request a native connection change for an external PivotTable.
Select both the worksheet and PivotTable
name: names can repeat on different worksheets. The target is an existing
workbook connection with the same OLEDB/ODBC type and OLAP mode, and must expose
compatible fields and measures. The action asks Excel to change the connection
in place rather than recreating the PivotTable. Excel may reject OLAP
connection changes even when the connection type and fields match; there is
no cache-rebuilding fallback.

```mcp
pivottable_read(action: 'get-connection', workbook_session_id: sessionId, sheet_name: 'Summary', pivot_table_name: 'SalesPivot')
pivottable(action: 'set-connection', workbook_session_id: sessionId, sheet_name: 'Summary', pivot_table_name: 'SalesPivot', connection_name: 'CurrentConnection')
```

```cli
excelcli -q pivottable get-connection --session $sessionId --sheet Summary --pivot-table-name SalesPivot
excelcli -q pivottable set-connection --session $sessionId --sheet Summary --pivot-table-name SalesPivot --connection-name CurrentConnection
```

The result reports the actual connection name and cache index, worksheet-qualified
shared-cache users, and connected slicer/timeline caches. Cache indexes can change
when Excel reorganizes caches; they are not permanent identifiers. Connection
strings and account details are not returned.

`set-connection` rejects shared caches and connected slicers/timelines rather
than changing other summaries or rebuilding controls. Disconnect controls
deliberately before changing their source. Worksheet-backed PivotTables, the
workbook's internal Data Model, and Power Query connection targets are not
converted by this action. Requesting the already-current connection is a no-op.

Changing a connection does not delete the old one or explicitly refresh the
PivotTable, although Excel may contact the provider during the change. Inspect
the returned state, then refresh and read the data. Only delete an old connection
after checking that no other PivotTables, worksheet queries, model tables, or
formulas still use it. Excel does not rewrite connection names embedded in
CUBE formulas or VBA.

Shared-cache settings can affect other summaries and therefore restrict edits.
After a failed write, inspect the actual state; do not assume rollback or
rebuild unrelated caches.
