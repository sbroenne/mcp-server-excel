# Charts

Chart lifecycle operations create, list, read, move, fit, and delete charts.
Chart-configuration operations manage series, titles, axes, labels, styles, and
trendlines. Reuse the returned chart name rather than assuming Excel's default.

## Choose the source before creating

Use the existing data directly when it already has useful categories and the
requested measures. Do not require a helper range, another Table, or extra
charts. Inspect the headers and a few rows first: numeric IDs, dates, percentages,
and amounts are not interchangeable series. A revenue comparison needs product
labels and revenue, not product IDs or every numeric column.

Choose a bounded range including its headers and intended data rows. Exclude
notes and grand totals that would duplicate the detail values. Keep categories
and values aligned, with the same row count. When a suitable Excel Table already
contains the requested fields, use Table-based creation to retain its source
behavior; check the plotted rows again after appending data. A fixed helper range
does not automatically grow with the original Table.

For a chart that must follow PivotTable fields and filters, use a live PivotChart,
not a regular chart of the displayed PivotTable cells. Change its series through
the [PivotTable fields](pivottable.md), not regular-chart series operations.

## Create and position

Choose the source deliberately: a range, an Excel Table, or a PivotTable. The
PivotTable action verifies a live PivotChart link; it fails rather than returning
a static chart if Excel cannot establish it.

For monthly labels in A1:A6 and numeric series in B1:C6 on `Sheet1`:

```mcp
chart(action: 'create-from-range', session_id: sessionId, sheet_name: 'Sheet1', source_range_address: 'A1:C6', chart_type: 'ColumnClustered', chart_name: 'MonthlySales', target_range: 'A8:H22')
chart_config(action: 'set-title', session_id: sessionId, chart_name: 'MonthlySales', title: 'Monthly sales')
chart(action: 'read', session_id: sessionId, chart_name: 'MonthlySales')
```

```cli
excelcli -q chart create-from-range --session $sessionId --sheet Sheet1 --source-range-address A1:C6 --chart-type ColumnClustered --chart-name MonthlySales --target-range A8:H22
excelcli -q chartconfig set-title --session $sessionId --chart-name MonthlySales --title 'Monthly sales'
excelcli -q chart read --session $sessionId --chart-name MonthlySales
```

Check each result. Prefer a target cell range for an exact placement. Omit both
target range and coordinates to auto-place below used cells and existing charts
with padding. Manual coordinates use points (72 per inch); row and column sizes
vary, so do not assume a fixed conversion from cells.

Creation, move, and fit operations warn about overlapping data/charts. An
`OVERLAP WARNING` can accompany success: fix the placement and check again.
Use a [screenshot](screenshot.md) when appearance matters and an interactive
desktop is available; otherwise inspect bounds and state the visual limitation.

## Percentage data labels

Percentage data labels are meaningful for pie and doughnut charts. MCP
`chart_config` action `set-data-labels` uses `show_percentage`; CLI
`excelcli chartconfig set-data-labels` uses `--show-percentage`. Excel can reject
the setting on other chart types; this is reported as an error, not a successful
no-op. Label settings are applied in sequence, so earlier settings can already
have changed when Excel rejects a later setting.

## Selected series, error bars and points

Use `get-series-settings` before changing a selected series. `set-series-chart-type`
changes its type; `set-series-axis-group` uses `axis_group` (MCP) /
`--axis-group` (CLI) to select Primary or Secondary. `chart` `read` also returns
each regular series' actual type and axis assignment. Axis titles, number
formats, scales and gridlines distinguish Category/Value from their Secondary
counterparts; select the intended axis explicitly.

`set-error-bars` takes `error_bar_options` / `--error-bar-options`, with camelCase
nested keys. Custom bars require both `plusRange` and `minusRange`: contiguous
one-dimensional numeric ranges with exactly one nonnegative value per point.
They are on `sourceSheetName`, or the chart worksheet when omitted. `direction`
is native Y (default) or X; X requires scatter/bubble. Y errors follow the value
axis, so they appear horizontal on a bar chart. `enabled:false` removes both
directions rather than clearing only the requested axis. Excel retains custom
source references only for the signs displayed by `include`.

`get-error-bars` reports native presence and caps. Excel provides no getters
for kind, direction, signs, amount or custom source references; `settingsReadable`
is false and `readLimitations` explains this. Do not treat previously sent
settings as native read-back.

`set-point-format` takes `point_index` / `--point-index` and `point_options` /
`--point-options`. Only the selected point changes. Marker points use their own
background/foreground colors, style and size; per-point marker transparency and
outline weight are rejected before mutation. Column-point transparency persists,
but Excel can still return an invalid transparency getter: `get-point-format`
reports `fillTransparencyAvailable:false` and an explicit read error, not a
fabricated zero. Automatic/mixed native color values are also explicitly
unavailable rather than converted into an invented RGB color.

These per-series writes reject PivotCharts because their fields/refresh control
the series. Native image export supports both regular charts and PivotCharts.
Use `chart` `export-image` with `target_path` / `--target-path`; extensions must
match the requested format and existing files require `overwrite:true` /
`--overwrite true`. The operation checks Excel's native result and nonempty
output, and does not silently replace an existing image on failure. Failed
exports remove their newly created output, including empty or partial images.

## Short labels without losing detail

Use this only when the existing labels are too long. Keep full descriptions and
source values intact. Put a small chart helper in a checked empty area, with
formulas linking its labels and amounts to the source rather than hardcoding a
second set of amounts. Include an identifier or another distinguishing part
when truncating text would produce duplicate labels. Grouping categories changes
the calculation, not just the label: state the grouping rule and summarize all
matching rows instead of keeping only one.

For existing `Details!A1:C4` containing `Product ID`, `Description`, and
`Revenue`, use empty `F1:G4` on the same sheet. The ID keeps shortened labels
distinct; the original description remains in column B:

```mcp
range(action: 'set-values', session_id: sessionId, sheet_name: 'Details', range_address: 'F1:G1', values: [['Product','Revenue']])
range(action: 'set-formulas', session_id: sessionId, sheet_name: 'Details', range_address: 'F2:G4', formulas: [['=A2&" - "&LEFT(B2,18)','=C2'],['=A3&" - "&LEFT(B3,18)','=C3'],['=A4&" - "&LEFT(B4,18)','=C4']])
calculation_mode(action: 'calculate', session_id: sessionId, scope: 'Sheet', sheet_name: 'Details')
range(action: 'get-formulas', session_id: sessionId, sheet_name: 'Details', range_address: 'F2:G4')
range(action: 'get-values', session_id: sessionId, sheet_name: 'Details', range_address: 'F1:G4')
chart(action: 'create-from-range', session_id: sessionId, sheet_name: 'Details', source_range_address: 'F1:G4', chart_type: 'BarClustered', chart_name: 'ProductRevenue')
```

```cli
excelcli -q range set-values --session $sessionId --sheet Details --range F1:G1 --values '[["Product","Revenue"]]'
excelcli -q range set-formulas --session $sessionId --sheet Details --range F2:G4 --formulas '[["=A2&\" - \"&LEFT(B2,18)","=C2"],["=A3&\" - \"&LEFT(B3,18)","=C3"],["=A4&\" - \"&LEFT(B4,18)","=C4"]]'
excelcli -q calculationmode calculate --session $sessionId --scope Sheet --sheet Details
excelcli -q range get-formulas --session $sessionId --sheet Details --range F2:G4
excelcli -q range get-values --session $sessionId --sheet Details --range F1:G4
excelcli -q chart create-from-range --session $sessionId --sheet Details --source-range-address F1:G4 --chart-type BarClustered --chart-name ProductRevenue
```

Check each result before continuing, including label uniqueness and unchanged
source amounts. These formulas follow fixed source rows: after a sort or source
growth, recheck the mapping and extend the helper and chart source as needed.
Do not silently replace an existing growing Table source with this fixed range.

## Readable periods in chronological order

Displaying a timestamp as a month does not aggregate the transactions. Use a
period summary when the requested trend needs monthly, quarterly, or yearly
values. Keep real dates as the period keys, display month/quarter/year labels,
and order by those dates rather than alphabetically by label. Keep the year:
December 2025 and January 2026 are consecutive periods, not a month-name sort.
State whether values are sums, counts, averages, or another measure; do not
average percentages without considering their denominators.

For an existing `Transactions` Table with native Excel `Date` values and numeric
`Revenue`, use an empty `Summary!A1:B3`. These two rows show monthly sums, including
timestamps on the last day and excluding the next month's first day:

```mcp
range(action: 'set-values', session_id: sessionId, sheet_name: 'Summary', range_address: 'A1:B3', values: [['Month','Revenue'],['2025-12-01',null],['2026-01-01',null]])
range(action: 'set-formulas', session_id: sessionId, sheet_name: 'Summary', range_address: 'B2:B3', formulas: [['=SUMIFS(Transactions[Revenue],Transactions[Date],">="&A2,Transactions[Date],"<"&EDATE(A2,1))'],['=SUMIFS(Transactions[Revenue],Transactions[Date],">="&A3,Transactions[Date],"<"&EDATE(A3,1))']])
range(action: 'set-number-format', session_id: sessionId, sheet_name: 'Summary', range_address: 'A2:A3', format_code: 'mmm yyyy')
calculation_mode(action: 'calculate', session_id: sessionId, scope: 'Sheet', sheet_name: 'Summary')
range(action: 'get-values', session_id: sessionId, sheet_name: 'Summary', range_address: 'A1:B3')
chart(action: 'create-from-range', session_id: sessionId, sheet_name: 'Summary', source_range_address: 'A1:B3', chart_type: 'Line', chart_name: 'MonthlyRevenue')
```

```cli
excelcli -q range set-values --session $sessionId --sheet Summary --range A1:B3 --values '[["Month","Revenue"],["2025-12-01",null],["2026-01-01",null]]'
excelcli -q range set-formulas --session $sessionId --sheet Summary --range B2:B3 --formulas '[["=SUMIFS(Transactions[Revenue],Transactions[Date],\">=\"&A2,Transactions[Date],\"<\"&EDATE(A2,1))"],["=SUMIFS(Transactions[Revenue],Transactions[Date],\">=\"&A3,Transactions[Date],\"<\"&EDATE(A3,1))"]]'
excelcli -q range set-number-format --session $sessionId --sheet Summary --range A2:A3 --format-code 'mmm yyyy'
excelcli -q calculationmode calculate --session $sessionId --scope Sheet --sheet Summary
excelcli -q range get-values --session $sessionId --sheet Summary --range A1:B3
excelcli -q chart create-from-range --session $sessionId --sheet Summary --source-range-address A1:B3 --chart-type Line --chart-name MonthlyRevenue
```

Reconcile representative monthly totals with the underlying transactions. For
quarters or years, use the actual period start and the next quarter/year start
as the boundaries. Include missing periods where the trend requires them;
`SUMIFS` returns zero with no matching rows, which is appropriate only when
absence means no activity, not unknown or incomplete data. Do not disguise
missing data as zero; choose the chart's blank-cell behavior deliberately.

For an interactive trend, summarize in the existing regular PivotTable instead.
For example, `RevenuePivot` already has native-date `Date` in Rows and Sum of
`Revenue` in Values. Month grouping also creates a year hierarchy:

```mcp
pivottable_field(action: 'group-by-date', session_id: sessionId, pivot_table_name: 'RevenuePivot', field_name: 'Date', interval: 'Months')
pivottable_field(action: 'list-fields', session_id: sessionId, pivot_table_name: 'RevenuePivot')
pivottable(action: 'refresh', session_id: sessionId, pivot_table_name: 'RevenuePivot')
pivottable_calc(action: 'get-data', session_id: sessionId, pivot_table_name: 'RevenuePivot')
chart(action: 'create-from-pivottable', session_id: sessionId, sheet_name: 'Summary', pivot_table_name: 'RevenuePivot', chart_type: 'Line', chart_name: 'InteractiveRevenue')
```

```cli
excelcli -q pivottablefield group-by-date --session $sessionId --pivot-table-name RevenuePivot --field-name Date --interval Months
excelcli -q pivottablefield list-fields --session $sessionId --pivot-table-name RevenuePivot
excelcli -q pivottable refresh --session $sessionId --pivot-table-name RevenuePivot
excelcli -q pivottablecalc get-data --session $sessionId --pivot-table-name RevenuePivot
excelcli -q chart create-from-pivottable --session $sessionId --sheet Summary --pivot-table-name RevenuePivot --chart-type Line --chart-name InteractiveRevenue
```

Inspect the generated fields rather than assuming their localized names. Check
the year/month order and totals before creating the chart; retain both year and
period distinctions. Date grouping requires valid dates without blank/error
items and is not supported for Data Model PivotTables. For those, use period
columns in the source/model and follow [PivotTable guidance](pivottable.md).

## Make units explicit

Choose display formats from the stored values, not from the column name alone.
Use an axis title or chart title to identify the currency and any scale.
Supply US format codes; Excel translates them for the user's locale, as with
[range number formats](range.md#number-formats-and-layout).
Axis formatting preserves explicit currency symbols and date/time meanings;
read-back returns US codes, not the localized COM codes.

| Stored meaning | Value-axis format | Important check |
|----------------|-------------------|-----------------|
| USD amounts | `$#,##0` | Do not assume every currency is USD |
| Fractional share, such as 0.35 | `0%` | Shows 35%; a stored 35 would show 3500% |
| Counts | `#,##0` | Do not relabel a monetary sum as a count |
| Unscaled USD amounts shown in thousands | `#,##0,` | 125000 displays as 125; title says USD thousands |
| Unscaled amounts shown in millions | `0.0,,` | Title identifies both unit and millions |

For the `MonthlyRevenue` chart above, keep the underlying amounts unchanged and
scale only its value-axis display:

```mcp
chart_config(action: 'set-axis-title', session_id: sessionId, chart_name: 'MonthlyRevenue', axis: 'Value', title: 'Revenue (USD thousands)')
chart_config(action: 'set-axis-number-format', session_id: sessionId, chart_name: 'MonthlyRevenue', axis: 'Value', number_format: '#,##0,')
chart_config(action: 'get-axis-number-format', session_id: sessionId, chart_name: 'MonthlyRevenue', axis: 'Value')
```

```cli
excelcli -q chartconfig set-axis-title --session $sessionId --chart-name MonthlyRevenue --axis Value --title 'Revenue (USD thousands)'
excelcli -q chartconfig set-axis-number-format --session $sessionId --chart-name MonthlyRevenue --axis Value --number-format '#,##0,'
excelcli -q chartconfig get-axis-number-format --session $sessionId --chart-name MonthlyRevenue --axis Value
```

Do not both divide helper values by 1000 and apply a thousands-scaling format.
If the source already stores thousands, use an ordinary numeric format and
label it accordingly. For whole-number percentages such as 35, use a clearly
identified formula-linked conversion to 0.35 if a percentage axis is needed;
do not change the original data silently.

Data-label options choose what to show, not a custom label number format.
Percentage labels on pie/doughnut charts mean each slice's share of the plotted
total; they are not a general percentage formatter for line/column charts.
For value labels, check their displayed units separately from the axis format;
an axis scaled to thousands does not establish that labels use the same scale.

## Configuration

Accepted option names are generated from the shared contracts into MCP parameter
descriptions and CLI help, including optional string-valued enum inputs.
Use those exact names rather than guessing Excel COM constant names or numeric
codes. The same values and case-insensitive validation apply to both entry points.

- Series indices are 1-based. Adding a series requires a values range; supply its
  category range when the axis labels are not implicit.
- Replacing the source range can change all series. Verify names, values, and
  categories afterward.
- Set per-series chart types for regular combo charts. Use plot options for
  row/column orientation, blanks, and whether hidden cells are plotted.
- Use Category and Value for primary axis titles. Use US number formats for
  currency/percentage tick labels; do not assume an axis format also formats
  data labels.
- CategorySecondary and ValueSecondary target the secondary axis group, which
  must already exist. The legacy Primary alias means Category, and Secondary
  means Value in the primary group; use the explicit names to avoid ambiguity.
- Built-in chart styles are 1-48. Area formatting controls chart/plot backgrounds;
  series formatting controls fills, lines, and markers.
- Placement 1 moves and sizes with cells, 2 moves only, 3 is free floating.
- Trendlines include Linear, Exponential, Logarithmic, Polynomial, Power, and
  MovingAverage. Polynomial order is 2-6; moving-average period is at least 2.
  Respect Excel's data/domain requirements for the selected fit.

For the existing `MonthlyRevenue` chart, use the owning chart controls:

```mcp
chart_config(action: 'show-legend', session_id: sessionId, chart_name: 'MonthlyRevenue', visible: true, legend_position: 'Bottom')
chart_config(action: 'set-series-format', session_id: sessionId, chart_name: 'MonthlyRevenue', series_index: 1, marker_style: 'Circle', marker_size: 5, line_color: '#0073BB')
chart_config(action: 'set-plot-options', session_id: sessionId, chart_name: 'MonthlyRevenue', plot_by: 'Columns', display_blanks_as: 'Gaps')
chart_config(action: 'set-area-format', session_id: sessionId, chart_name: 'MonthlyRevenue', area: 'Plot', fill_color: '#FFFFFF')
chart(action: 'read', session_id: sessionId, chart_name: 'MonthlyRevenue')
```

```cli
excelcli -q chartconfig show-legend --session $sessionId --chart-name MonthlyRevenue --visible true --legend-position Bottom
excelcli -q chartconfig set-series-format --session $sessionId --chart-name MonthlyRevenue --series-index 1 --marker-style Circle --marker-size 5 --line-color '#0073BB'
excelcli -q chartconfig set-plot-options --session $sessionId --chart-name MonthlyRevenue --plot-by Columns --display-blanks-as Gaps
excelcli -q chartconfig set-area-format --session $sessionId --chart-name MonthlyRevenue --area Plot --fill-color '#FFFFFF'
excelcli -q chart read --session $sessionId --chart-name MonthlyRevenue
```

Check each result and the chart's actual data. Formatting a PivotChart does not
authorize changing its field configuration or filter scope.

For multiple charts, use explicit non-overlapping cell ranges with consistent
sizes and spacing. Auto-placement is suitable for a vertical stack. Read the
saved chart's actual geometry and series; a successful creation or a prose
description alone does not establish a correct chart.

## Check data, not just appearance

Read the chart after creation or a source change. Check its series names/count
and available source information, then read the referenced cells and helper
formulas or PivotTable totals. Compare representative amounts with the original
rows; check category/value alignment, period order, and the treatment of totals,
hidden rows, blanks, and errors.

For both regular charts and PivotCharts, `series` includes the actual plotted
names, `values`, and `categories` in point order. The listed `seriesCount` counts
plotted series, not PivotTable value fields: one revenue measure can produce
three provider series. Filtering the linked pivot changes those plotted series.

For regular charts, `sourceRange` still contains the first series' `SERIES`
formula, not a complete source rectangle. On read, `valuesRange` is empty and
`categoryRange` is null rather than inventing cell addresses from Excel's value
arrays. Series-creation operations can return the supplied ranges separately.
Use the plotted arrays for data checks; a formula alone does not establish every
binding. PivotCharts also report `isPivotChart` and `linkedPivotTable`; verify
that link and the actual PivotTable fields, filters, and data.

Keep checks and changes within the request. Changing a title or unit display
does not authorize replacing the source, rebuilding unrelated data, restyling
the workbook, adding charts, or repairing unrelated pre-existing errors.
