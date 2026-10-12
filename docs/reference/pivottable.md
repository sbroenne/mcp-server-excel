# PivotTables

Reuse a suitable existing PivotTable. Creating one does not configure its
rows, columns, values, or filters; choose those fields, then refresh and read
the actual summary.

Current commands and inputs come from CLI help or MCP tool descriptions.

| Source | Calculation approach |
|--------|----------------------|
| Worksheet Table or range | Ordinary aggregations and regular calculated fields |
| Data Model | DAX measures and model relationships |
| External OLAP cube | Existing server measures and cube hierarchies; supported client calculated members use MDX |

## Read a bounded configuration snapshot

`pivottable` action `read` returns the PivotTable layout together with exact
selected members, native label/value filters, Values-field settings, and
whether Excel is deferring updates. It does not return report cell values,
refresh the PivotTable, or save the workbook.

Use MCP `sheet_name` (CLI `--sheet`) to choose the exact worksheet when
PivotTable names may repeat. MCP `max_selected_members` (CLI
`--max-selected-members`) sets the maximum number of selected member names
returned across all fields; it defaults to 100 and accepts 1 through 1000.
Names are sorted for stable comparisons. `selectedMemberCount` reports the
full count when known, while `membersTruncated` and `snapshotComplete` make an
incomplete result explicit.

For each placed field, `fieldSelections` identifies the semantic field name and
area, reports whether its exact selection and native filters were available,
and lists selected members when `showAll` is false. An All selection has
`showAll: true` and no selected member list. `fieldSelectionsAvailable`,
`valueFieldsAvailable`, and `isOlapAvailable` distinguish unavailable
configuration getters from empty results; inspect their matching error fields
instead of assuming an empty list means no fields or selections.
`snapshotComplete` is false when required state could not be read or the member
bound truncated the result.

```mcp
pivottable(action: 'read', workbook_session_id: sessionId, pivot_table_name: 'SalesPivot', sheet_name: 'Summary', max_selected_members: 50)
```

```cli
excelcli -q pivottable read --session $sessionId --pivot-table-name SalesPivot --sheet Summary --max-selected-members 50
```

For an existing cube measure, `add-value-field` accepts its exact qualified
name (`[Measures].[Amount.SUM]`) or bare name (`Amount.SUM`). Dots in a
measure name are not hierarchy separators. The requested aggregation is ignored
because the source measure defines it. Only workbook Data Model columns can
automatically create DAX measures.
Closing brackets use MDX escaping in qualified names: bare `Margin]Pct` is
`[Measures].[Margin]]Pct]`.

`set-field-function` changes only a single-column aggregation measure in the
workbook Data Model. It must not replace a compound formula or change an
unrelated workbook measure with the same name as a server measure.
Use `datamodel update-measure` for an intentional compound DAX formula change.
External measures must be changed on the cube server.

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

Excel does not expose every model-backed base-field/item setting. An OLAP/Data
Model calculation write can change displayed values while the native getter
still reports the previous calculation. A readback mismatch returns a verification
failure, not proof that the write was ignored or rolled back. Inspect `get-data`
as well as `list-fields` before retrying or explicitly changing the calculation.
Use a source measure when native settings cannot establish the result.

Calculated-member creation reports formula syntax only when Excel's diagnostic
explicitly identifies syntax. A generic Excel or provider failure is not proof
of an invalid formula. After a failure, inspect `list-calculated-members` before
retrying; the operation does not promise rollback.
For a created measure, run `pivottable refresh` to expose it in CubeFields before
`add-value-field`. Use its qualified name without adding
another `[Measures]` prefix (`pivottable_field` input `field_name`, CLI
`pivottablefield --field-name`). Created members and sets are not value fields;
inspect them with `list-calculated-members` instead.

## Refresh, layout, and grouping

Hierarchy subtotal changes apply one level at a time. A provider failure reports
the completed levels and the attempted level, whose state may also have changed.
Nothing is rolled back; inspect the hierarchy before retrying.

Refresh sources before their summaries. Worksheet-source edits may need a model
refresh, followed by PivotTable refresh; Power Query refresh updates its model
load but does not replace the dependent PivotTable refresh.

- `sort-field` orders the selected field's labels in ascending or descending
  order, including Data Model fields identified by their exact CubeField name.
  To sort by values instead, supply MCP `data_field_name` or CLI
  `--data-field-name`, using the exact `valueFields.fieldName` returned by
  `list-fields`. This distinguishes repeated Values instances (for example Sum
  and Average of Sales). OLAP sorting uses the selected measure on the
  hierarchy's first placed level.
- `set-field-name` updates a regular PivotTable's displayed value-field caption when the selected
  source field is in Values; it does not rename the source worksheet column.
  Use MCP `field_name` / `custom_name` or CLI `--field-name` / `--custom-name`.
  OLAP hierarchy/measure names come from the Data Model or external cube;
  this action does not rename them.
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
For external cubes, use server-defined hierarchies and grouping fields.
`set-subtotals` accepts the exact OLAP CubeField hierarchy name and changes
automatic subtotals on its levels; the hierarchy must be placed in Rows or Columns.
Drill-through on a regular value cell creates a new sheet of underlying rows;
it is not a read-only audit.

Use a verified live [PivotChart](chart.md) when a chart must follow fields and
filters, rather than a regular chart of displayed cells.

## Additional reporting controls

`get-field-options` and `set-field-options` read/change whether a regular
row or column field includes items with no data (MCP `show_all_items`, CLI
`--show-all-items`). These show cached items, not categories absent from the
source. OLAP/Data Model uses `set-layout-options` with nested
`layoutOptions.showEmptyRows` and `layoutOptions.showEmptyColumns` instead;
the provider determines which empty members are available.

For regular row/column fields, `get-subtotal-functions` returns the complete
selection and `set-subtotal-functions` replaces it. An empty `functions` array
removes subtotals; `Automatic` must stand alone. MCP takes an actual array,
such as `functions: ["Sum","Average"]`; CLI uses a JSON array string:
`--functions '["Sum","Average"]'`. OLAP/Data Model continues to use
automatic `set-subtotals`; custom subtotal functions are not supported there.

`create-calculated-item`, `list-calculated-items`, and `delete-calculated-item`
operate on items inside a regular row/column field, not numeric source columns
or Values instances. For example, in a Region field with North and South,
`item_name: 'Combined', formula: '=North+South'` (CLI `--item-name Combined
--formula '=North+South'`) adds a new Combined row. That row participates in
grand totals and may double-count the underlying sales. Creation requires a
new item name. Deletion accepts only an existing calculated item, never a source
item. Calculated-item writes reject PivotTables with grouped fields; OLAP/Data
Model calculations must use source measures or supported calculated members.

`set-layout-options` also accepts the following nested `layoutOptions` keys
(MCP top-level `layout_options`, CLI `--layout-options`):

- `displayNullString` / `nullString` and `displayErrorString` / `errorString`
  control custom blank/error text in Values cells. Text may be empty. Set the
  corresponding display flag explicitly to enable or disable replacement;
  this is presentation, not a change to source values.
- `mergeLabels`, `printTitles`, and `repeatItemsOnEachPrintedPage` control
  merged labels and repeated printed headings/items for the whole PivotTable.
- `blankLineAfterEachItem` applies blank separators to every placed row field.
  `get-layout-options` returns each row field's actual separator setting.

Omitted options leave their settings unchanged. Read the returned native
settings and `get-data` to check the result. Failures do not promise rollback;
inspect the affected report before retrying or deliberately undoing a change.

## Native filters, layout, and source changes

Manual item visibility and calculated filters are different. Clearing calculated
filters need not clear manual selections or other fields. Use the actual placed
field and displayed Values captions rather than guessing from source names.
Multiple filters are not automatically authorized by adding another criterion.

For an OLAP/Data Model report filter, use `set-report-filter` rather than
`set-field-filter`. Target the worksheet and PivotTable together, and pass exact
OLAP member unique names (for example,
`[RegionTable].[Region].&[North]`). An empty `selected_items` list shows all
members; one or several names select those members only. This changes only the
chosen PivotTable's report filter, not the Data Model or other reports. Use the
exact hierarchy unique name returned by `list-fields` for `field_name`, not a
level-qualified name. A one-member selection preserves the field's existing
single- or multiple-selection mode; selecting several members enables Excel's
multiple-selection mode when needed.
After reopening a workbook with an external OLAP source, Excel may require an
explicit PivotTable refresh to reconnect before the filter can be changed;
refreshing can update source data, so it is not done automatically.
For All, pass `selected_items: '[]'` in MCP or `--selected-items '[]'` in the CLI.

Use `get-report-filter` with the same worksheet, PivotTable, and hierarchy to
read one report filter without returning every PivotTable field or Values
instance. It returns the exact selected member unique names, All state,
selection mode, and field placement. `selectionStateAvailable: false` means
Excel could not read the state; do not interpret empty `selectedItems` as All
unless `showAll` is true and the state is available.

```mcp
pivottable_field(action: 'set-report-filter', workbook_session_id: sessionId, sheet_name: 'Summary', pivot_table_name: 'SalesPivot', field_name: '[RegionTable].[Region]', selected_items: '["[RegionTable].[Region].&[North]","[RegionTable].[Region].&[West]"]')
pivottable_field(action: 'get-report-filter', workbook_session_id: sessionId, sheet_name: 'Summary', pivot_table_name: 'SalesPivot', field_name: '[RegionTable].[Region]')
```

```cli
excelcli -q pivottablefield set-report-filter --session $sessionId --sheet Summary --pivot-table-name SalesPivot --field-name '[RegionTable].[Region]' --selected-items '["[RegionTable].[Region].&[North]","[RegionTable].[Region].&[West]"]'
excelcli -q pivottablefield get-report-filter --session $sessionId --sheet Summary --pivot-table-name SalesPivot --field-name '[RegionTable].[Region]'
```

The write result separates `requestedItems` from `selectedItems`, which reports
the actual state when Excel allowed readback. `failurePhase` identifies the
failing operation phase. If `mayHavePartiallyChanged` is true after a failure,
inspect the filter with `get-report-filter` before retrying. The result also
reports `rollbackAttempted`; this operation does not automatically undo a
failed filter change.
The operation defers updates on the selected PivotTable where supported, but
does not guarantee a particular number of provider queries.

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
