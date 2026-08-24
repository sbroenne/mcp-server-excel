# Build and Update PivotTables with an AI Assistant

PivotTables are the part of Excel most people want an assistant to handle, and the
part that file-parser libraries handle worst. ExcelMcp creates and refreshes them
through Excel's own PivotTable engine, so the result is a real, interactive
PivotTable — not a static grid of pre-aggregated numbers.

## What you ask for

> In `sales.xlsx`, build a PivotTable of revenue by region and quarter, then chart it.

> The `SalesTable` grew — refresh `SalesPivot` and tell me the new total.

## Pick the right source first

This decision determines what you can do later, so make it deliberately:

| Source | Create action | DAX measures? |
|---|---|---|
| Worksheet Excel Table | `create-from-table` | No |
| Data Model (Power Pivot) | `create-from-datamodel` | **Yes** |
| Range on a sheet | `create` with a source range | No |

**Rule of thumb:** if the analysis needs calculated values, multiple tables, or
reusable measures, build it on the Data Model. Choosing a worksheet table and
discovering later that DAX measures are missing means rebuilding.

## Build one from a worksheet table

```powershell
$session = (excelcli -q session open C:\data\sales.xlsx | ConvertFrom-Json).sessionId

excelcli -q pivottable create-from-table --session $session `
  --table-name SalesTable --pivot-table-name SalesPivot `
  --destination-sheet Analysis --destination-cell A3

excelcli -q pivottablefield add-row-field    --session $session --pivot-table-name SalesPivot --field-name Region
excelcli -q pivottablefield add-column-field --session $session --pivot-table-name SalesPivot --field-name Quarter
excelcli -q pivottablefield add-value-field  --session $session --pivot-table-name SalesPivot --field-name Amount --aggregation-function Sum

excelcli -q pivottable refresh --session $session --pivot-table-name SalesPivot
excelcli -q session close --session $session --save
```

Configure fields in that order — rows, columns, values, then filters — and
**always finish with `refresh`**. Field operations are structural only: they change
the layout but do not repaint the PivotTable. This matters most for Data Model
(OLAP) PivotTables, where an unrefreshed pivot can look empty.

`--pivot-table-name` is required for nearly every PivotTable operation. The only
exception is `list`.

## Build one on the Data Model

This is the path that supports DAX, multiple tables, and reusable measures:

```powershell
excelcli -q table add-to-data-model --session $session --table-name SalesTable

excelcli -q datamodel create-measure --session $session `
  --table-name SalesTable --measure-name Revenue `
  --dax-formula "SUMX(SalesTable, SalesTable[Quantity]*SalesTable[UnitPrice])"

excelcli -q pivottable create-from-datamodel --session $session `
  --pivot-table-name RevenuePivot --table-name SalesTable `
  --destination-sheet Analysis --destination-cell A3
```

The measure is automatically available to the PivotTable — no calculated field
needed.

### Calculated field or DAX measure?

| | PivotTable calculated field | DAX measure |
|---|---|---|
| Single-table formula (`=Qty*Price`) | Works | Works |
| Across related tables | Not supported | Works |
| Time intelligence, YTD, running totals | Limited | Works |
| Reusable elsewhere | That PivotTable only | Whole Data Model |

Use a calculated field for something trivial and local. Use DAX for anything else.

## Chart it in one step

Do **not** create a PivotTable and then a separate chart from its cells. Create a
PivotChart directly — it is one object bound to the same cache:

```powershell
excelcli -q chart create-from-pivottable --session $session --sheet Analysis `
  --pivot-table-name SalesPivot --chart-type ColumnClustered
```

## Verify

```powershell
excelcli -q pivottable list --session $session
excelcli -q screenshot capture-sheet --session $session --sheet Analysis
```

A screenshot is the fastest way for an assistant to confirm a layout actually looks
right, rather than inferring it from return values.

## Known gotchas

**PivotTables never auto-refresh.** Changing source data does nothing until you
call `refresh`. After appending rows to a worksheet table you need *two* refreshes
if DAX is involved:

```text
table(action: 'append', ...)     # worksheet table updated
datamodel(action: 'refresh')     # sync the Data Model copy
pivottable(action: 'refresh')    # repaint the PivotTable
```

After a Power Query refresh this is handled for you — Power Query refreshes the
Data Model, and Data Model PivotTables follow.

**Custom formatting does not survive a refresh.** Colours, bold, and borders
applied to PivotTable cells are erased whenever Excel's layout engine reapplies
defaults. Use the PivotTable's own style (`set-style`) instead of formatting
individual cells — styles persist.

**"Unknown field" on a value field** usually means a calculated-field limitation.
Switch to a DAX measure.

**"Table not found"** means the source was never added to the Data Model. Run
`table add-to-data-model` first.

**Manual grouping needs a regular PivotTable** with the field already placed in the
row or column area. Data Model PivotTables must add the grouping column in the
model instead.

**Layout style** is worth setting explicitly: compact (default, nested labels),
tabular (one field per column — best for exports), or outline.

## Related

- [Charts, PivotTables and visual operations](../features/CHARTS-VISUALS.md)
- [Query the Data Model with DAX](QUERY-DATA-MODEL-WITH-DAX.md)
- [Refresh Power Query from an AI assistant](REFRESH-POWER-QUERY.md)
