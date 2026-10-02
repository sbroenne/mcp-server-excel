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

## Refresh, layout, and grouping

- After source edits, refresh the Data Model if used, then the PivotTable.
  Power Query refresh updates its model load; refresh the PivotTable afterward.
- After field changes, refresh once all fields are configured, especially for
  model-backed PivotTables.
- Layout values are 0 Compact, 1 Tabular, 2 Outline. Tabular is useful for exports.
- Use field number formats, not plain-range visual formatting on PivotTable cells.
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
