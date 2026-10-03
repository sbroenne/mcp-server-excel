# PivotTables

List before creating. Most operations require the PivotTable name; creating a
PivotTable does not configure its row, column, or value fields. Add those fields
with the field operations, then refresh and read the actual data.

| Source | Creation | Calculation support |
|--------|----------|---------------------|
| Worksheet Table | `create-from-table` | Ordinary aggregations and calculated fields |
| Worksheet range | `create-from-range` | Ordinary aggregations and calculated fields |
| Data Model | `create-from-datamodel` | DAX measures and model relationships |

## Calculated fields are calculations on aggregates

For a regular PivotTable with a numeric `Sales` field, doubling Sales is safe:

```mcp
pivottable_calc(action: 'create-calculated-field', session_id: sessionId, pivot_table_name: 'SalesPivot', field_name: 'DoubleSales', formula: '=Sales*2')
pivottable_field(action: 'add-value-field', session_id: sessionId, pivot_table_name: 'SalesPivot', field_name: 'DoubleSales', aggregation_function: 'Sum')
pivottable(action: 'refresh', session_id: sessionId, pivot_table_name: 'SalesPivot')
pivottable_calc(action: 'get-data', session_id: sessionId, pivot_table_name: 'SalesPivot')
```

```cli
excelcli -q pivottablecalc create-calculated-field --session $sessionId --pivot-table-name SalesPivot --field-name DoubleSales --formula '=Sales*2'
excelcli -q pivottablefield add-value-field --session $sessionId --pivot-table-name SalesPivot --field-name DoubleSales --aggregation-function Sum
excelcli -q pivottable refresh --session $sessionId --pivot-table-name SalesPivot
excelcli -q pivottablecalc get-data --session $sessionId --pivot-table-name SalesPivot
```

Check each result before the next step. Creation alone does not add the
calculated field to Values. Excel may report its raw field type as text; the
server identifies calculated fields as numeric for aggregation.

Do **not** use `=Quantity*UnitPrice` as a PivotTable calculated field for
line-item revenue. For rows `(2,10)` and `(3,20)`, the sum of row revenues is 80,
not the product of summed inputs (150). Add a source Revenue column with a formula
on every data row and sum that column, or create this Data Model measure:

```dax
SUMX(Sales, Sales[Quantity] * Sales[UnitPrice])
```

DAX requires the table in the model first. Use
[Data Model guidance](datamodel.md) for relationships, measures, and display.

## Show Values As is separate from aggregation

Use `pivottable_field` (MCP) / `pivottablefield` (CLI) `list-fields` and select
the exact `fieldName` from `valueFields`. Each displayed Values instance has
its own additional calculation, even when several instances share a source.
Do not select the source name to guess which instance should change.
`set-field-calculation` changes only that instance, not its Sum/Average function.
`Normal` removes the additional calculation without changing aggregation.

```mcp
pivottable_field(action: 'set-field-calculation', session_id: sessionId, pivot_table_name: 'SalesPivot', field_name: 'Total Sales', calculation: 'PercentOfTotal')
pivottable_calc(action: 'get-data', session_id: sessionId, pivot_table_name: 'SalesPivot')
pivottable_field(action: 'set-field-calculation', session_id: sessionId, pivot_table_name: 'SalesPivot', field_name: 'Total Sales', calculation: 'DifferenceFrom', base_field_name: 'Region', base_item_kind: 'Previous')
```

```cli
excelcli -q pivottablefield set-field-calculation --session $sessionId --pivot-table-name SalesPivot --field-name "Total Sales" --calculation PercentOfTotal
excelcli -q pivottablecalc get-data --session $sessionId --pivot-table-name SalesPivot
excelcli -q pivottablefield set-field-calculation --session $sessionId --pivot-table-name SalesPivot --field-name "Total Sales" --calculation DifferenceFrom --base-field-name Region --base-item-kind Previous
```

Base-dependent calculations require an exact field placed in Rows or Columns.
Named base items use exact item names, never numeric indexes. Relative items
follow Excel's current item order; sorting/filtering can change their results.
Parent percentages require the relevant row/column axis. Native blank/error
results are not converted into zero, and failed writes do not promise rollback.
Excel does not expose base-field/item settings for OLAP/Data Model PivotTables;
use a source measure for base-dependent calculations. Excel can also silently
ignore additional calculations on model measures. Every write verifies native
settings and reports failure if Excel did not apply them; use a source measure
in that case. Read back both settings and displayed data before relying on results.

## Refresh, layout, and grouping

- After source edits, refresh the Data Model if used, then the PivotTable.
  Power Query refresh updates its model load; refresh the PivotTable afterward.
- After field changes, refresh once all fields are configured, especially for
  model-backed PivotTables.
- `sort-field` orders the selected field's labels in ascending or descending
  order, including Data Model fields identified by their exact CubeField name.
- Layout values are 0 Compact, 1 Tabular, 2 Outline. Tabular is useful for exports.
- Use field number formats, not plain-range visual formatting on PivotTable cells.
- `set-field-name` updates the displayed value-field caption when the selected
  source field is in Values; it does not rename the source worksheet column.
  Use MCP `field_name` / `custom_name` or CLI `--field-name` / `--custom-name`.
- `set-field-format` accepts invariant Excel number formats. A bare dollar sign
  remains a literal dollar sign, not the machine's regional currency. Escaped or
  quoted literals and bracketed currency/locale codes remain intact. Excel may
  return a normalized format with an escaped dollar sign. Use MCP `number_format`
  or CLI `--number-format`.
- Manual grouping needs a regular PivotTable field in Rows or Columns. Grouped
  field names come from the result; use that returned name to ungroup.
- Date/numeric grouping and calculated fields are not model-backed operations;
  add grouping columns or measures in the source/model instead.
- Drill-through requires a regular PivotTable value cell and creates a worksheet
  of underlying rows. The provider-dependent model equivalent is not exposed.
- Cache options control refresh, retained items, and saved source data. Retained
  item limits apply to regular caches, not model-managed members.

For charts that must follow field/filter changes, create a verified live
[PivotChart](chart.md), not a static chart of the displayed cells.

## Native filters, layout, and source changes

`set-field-filter` controls manual item visibility. The separate
`get-field-filters` / `add-field-filter` / `clear-field-filters` actions on
`pivottable_field` (MCP) / `pivottablefield` (CLI) control calculated filters.
Clearing calculated filters preserves other fields and manual item selections.
Select the exact placed field caption; value/top filters also require the exact
displayed Values caption in `dataFieldName`, not its source name.
Date criteria use ISO dates in `date1` / `date2`; native reads return the actual
criteria, which Excel can expose as numeric date serials.

```mcp
pivottable_field(action: 'add-field-filter', session_id: sessionId, pivot_table_name: 'SalesPivot', field_name: 'Region', filter_options: { type: 'ValueIsGreaterThan', number1: 150, dataFieldName: 'Total Sales' })
pivottable_calc(action: 'set-layout-options', session_id: sessionId, pivot_table_name: 'SalesPivot', layout_options: { rowLayout: 1, repeatLabels: true, styleName: 'PivotStyleMedium9', preserveFormatting: true })
```

```cli
excelcli -q pivottablefield add-field-filter --session $sessionId --pivot-table-name SalesPivot --field-name Region --filter-options '{"type":"ValueIsGreaterThan","number1":150,"dataFieldName":"Total Sales"}'
excelcli -q pivottablecalc set-layout-options --session $sessionId --pivot-table-name SalesPivot --layout-options '{"rowLayout":1,"repeatLabels":true,"styleName":"PivotStyleMedium9","preserveFormatting":true}'
```

An existing calculated filter is not silently replaced when multiple filters
are disabled. Clear it intentionally or set `layoutOptions.allowMultipleFilters`
to true first. `get-layout-options` reads every row field separately, including
mixed layout/repetition. Repeated labels require Tabular or Outline layout.
Use native PivotTable styles rather than restyling its cells after every refresh.
`get-item-expansion` / `set-item-expansion` target one exact visible parent item,
not all members; innermost fields have no child level.
Calculated-filter and item-expansion operations reject OLAP/Data Model tables
because their provider-dependent behavior is not supported here.

Use `pivottable` `get-source` before an intended source change. `set-source`
accepts `source_sheet_name` and exactly one `source_range_address` / `table_name`
in MCP, or `--source-sheet-name` and exactly one `--source-range-address` /
`--table-name` in CLI. The replacement must retain the same source field names.
It creates a new cache only for the selected worksheet-backed PivotTable,
refreshes it, and leaves unrelated cache users unchanged. Connected slicers or
timelines must first be explicitly disconnected; their caches are not rebuilt.
External/OLAP source changes are rejected. Cache-wide `set-cache-options`
changes are rejected when other PivotTables share the cache; saving source data
remains table-specific. Failed writes do not promise rollback.
