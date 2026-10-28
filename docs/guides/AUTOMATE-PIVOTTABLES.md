# Build and Update PivotTables with an AI Assistant

ExcelMcp creates real, interactive PivotTables through desktop Excel. Inspect
existing reports with `pivottable list` before creating another; refresh or
modify a suitable report rather than losing its configuration by recreating it.

## Choose the source

| Source | Create action | Use for |
|---|---|---|
| Worksheet Excel Table | `create-from-table` | A source that expands as rows are added |
| Fixed worksheet range | `create-from-range` | An explicitly bounded source |
| Data Model | `create-from-datamodel` | DAX measures and related tables |

Creation does not place fields. Add rows, columns, values and filters as needed,
then refresh and inspect `pivottablecalc get-data`. Use `pivottable read` for
configuration, field selections and `cacheIndex`, not report cells.

## Two reports, one slicer

This PowerShell example creates a new workbook with synthetic sales, two reports
sharing one cache, and a Region slicer connected to both. Use an existing
directory and a new workbook path. Check each command's success response before
continuing; a failed write does not promise rollback.

```powershell
$session = (excelcli -q session create C:\data\dashboard.xlsx | ConvertFrom-Json).data.sessionId
excelcli -q sheet create --session $session --sheet-name Data
excelcli -q sheet create --session $session --sheet-name Analysis
excelcli -q range set-values --session $session --sheet-name Data --range-address A1:C4 `
  --values '[["Region","Product","Amount"],["North","Widget",100],["South","Widget",200],["North","Gadget",50]]'
excelcli -q table create --session $session --sheet-name Data --table-name SalesTable --range-address A1:C4

excelcli -q pivottable create-from-table --session $session `
  --table-name SalesTable --pivot-table-name ByRegion `
  --destination-sheet Analysis --destination-cell A3
excelcli -q pivottable create-from-table --session $session `
  --table-name SalesTable --pivot-table-name ByProduct `
  --destination-sheet Analysis --destination-cell H3 `
  --share-cache-with ByRegion --share-cache-with-sheet Analysis

excelcli -q pivottablefield add-row-field --session $session --pivot-table-name ByRegion --field-name Region
excelcli -q pivottablefield add-row-field --session $session --pivot-table-name ByProduct --field-name Product
foreach ($pivot in 'ByRegion', 'ByProduct') {
  excelcli -q pivottablefield add-value-field --session $session `
    --pivot-table-name $pivot --field-name Amount --aggregation-function Sum
  excelcli -q pivottable refresh --session $session --pivot-table-name $pivot
}

excelcli -q slicer create-slicer --session $session --pivot-table-name ByRegion `
  --field-name Region --slicer-name RegionFilter --destination-sheet Analysis --position O3
excelcli -q slicer connect-pivottable --session $session --slicer-name RegionFilter --pivot-table-name ByProduct
excelcli -q slicer set-slicer-selection --session $session --slicer-name RegionFilter --selected-items '["North"]'
excelcli -q pivottablecalc get-data --session $session --pivot-table-name ByRegion
excelcli -q pivottablecalc get-data --session $session --pivot-table-name ByProduct
# Both totals are 150; ByProduct shows Widget=100 and Gadget=50.
excelcli -q slicer set-slicer-selection --session $session --slicer-name RegionFilter --selected-items '[]'
excelcli -q pivottablecalc get-data --session $session --pivot-table-name ByRegion
excelcli -q pivottablecalc get-data --session $session --pivot-table-name ByProduct
# Clearing restores both totals to 350.
excelcli -q pivottable list --session $session
excelcli -q session close --session $session --save
```

The sharing pair also works with `create-from-range`. Omit both options for
independent caches. Identical data alone is not sufficient: sharing requires the
same source identity, and Table-versus-fixed-range semantics stay distinct.
Creation options do not move existing reports onto another cache. See the
[PivotTable reference](../reference/pivottable.md#create-another-report-sharing-a-worksheet-cache)
for matching MCP examples, validation and shared-cache consequences.

For an already-built `ByProduct` report with an independent cache, use
`excelcli -q pivottable share-cache --session $session --sheet-name Analysis --pivot-table-name ByProduct --share-cache-with ByRegion --share-cache-with-sheet Analysis`
instead of recreating it. The reports must have matching worksheet sources and
compatible settings. Disconnect controls from the moved report first, then
connect a shared slicer explicitly afterward. See
[existing report sharing](../reference/pivottable.md#move-an-existing-report-onto-a-shared-cache)
for preservation limits and failure recovery.

## Calculations, refresh and presentation

- **Calculations:** regular calculated fields operate on aggregates, not source
  rows, and their definitions are shared by cache. Use a source revenue column
  or a DAX measure for row-wise quantity × price. The
  [calculation limitations](../reference/pivottable.md#calculated-fields-are-calculations-on-aggregates)
  explain the distinction.
- **Refresh:** update the source first, then refresh dependent PivotTables.
  Power Query refresh does not replace that PivotTable refresh. Shared cache
  refresh and native grouping can affect other reports; see
  [refresh and grouping](../reference/pivottable.md#refresh-layout-and-grouping).
- **Layout:** use `pivottablecalc set-layout-options` with
  `--layout-options '{"rowLayout":1,"styleName":"PivotStyleMedium9","preserveFormatting":true}'`
  and `--pivot-table-name` / `--session` for the chosen report. Native styles and
  formatting preservation are preferable to relying on plain-cell formatting
  through refresh.
- **Charts:** `chart create-from-pivottable` creates a linked PivotChart when a
  chart should follow report fields and filters. A screenshot is optional when
  an interactive desktop is available, not required for verification.

## Related

- [PivotTable workflows and limitations](../reference/pivottable.md)
- [Charts, PivotTables and visual operations](../features/CHARTS-VISUALS.md)
- [Query the Data Model with DAX](QUERY-DATA-MODEL-WITH-DAX.md)
- [Refresh Power Query from an AI assistant](REFRESH-POWER-QUERY.md)
