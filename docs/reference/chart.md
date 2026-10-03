# Charts

Build the chart from the intended data, then check its actual plotted results.
Reuse the returned chart name rather than assuming Excel's default. Current
commands and inputs come from CLI help or MCP tool descriptions.

For `chart_config` / `chartconfig` `set-data-labels`, line and scatter series
do not support `InsideEnd`, `InsideBase`, or `OutsideEnd`. Those requests are
rejected before any targeted series changes, including combination charts.
Use `Above`, `Below`, `Left`, `Right`, or `Center` for line and scatter series.

For `chart_config` / `chartconfig` `add-series`, `values_range` / `--values-range`
and an optional `category_range` / `--category-range` must resolve to existing
Excel ranges. A missing sheet or invalid range is rejected before a series is
added, preserving the existing chart. The added series retains live references
to those ranges; changes to their cells update the plotted values.

Axis number formats use US codes, including named colors such as `[Red]`.
Excel's chart format properties have different regional rules from worksheet
cells; the server handles that difference without changing regional settings.
Reads return canonical codes and may add literal-dollar escapes.

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

Check each result. Prefer a target cell range for an exact placement, or use
supported automatic placement when a vertical stack suits the report.
Manual coordinates use points; row and column sizes
vary, so do not assume a fixed conversion from cells.

Creation, move, and fit operations warn about overlapping data/charts. An
`OVERLAP WARNING` can accompany success: fix the placement and check again.
Use a [screenshot](screenshot.md) when appearance matters and an interactive
desktop is available; otherwise inspect bounds and state the visual limitation.

## Percentage data labels

Percentage data labels are meaningful for pie and doughnut charts. Excel can reject
the setting on other chart types; this is reported as an error, not a successful
no-op. Label settings are applied in sequence, so earlier settings can already
have changed when Excel rejects a later setting.

## Selected series, error bars and points

Inspect the intended series before changing its type or axis assignment.
Primary and secondary axes can use different scales; identify the actual axis
instead of treating similarly named selectors as interchangeable.

Custom error bars need numeric source values aligned with every plotted point.
Their orientation follows the native axes, which can differ from the visual
direction suggested by a bar chart. Turning error bars off removes both
directions, not just a selected direction.

Excel cannot read back every error-bar setting it accepts. Previously sent
settings are not native readback. The same caution applies to some point
transparency and automatic/mixed colors: explicit inspection limitations are
not zero values or proof that formatting failed.

Selected-point formatting affects only that point, but supported properties
depend on whether the point uses markers or a material fill. These per-series
writes reject PivotCharts because their fields and refresh control the series.

Native image export supports regular charts and PivotCharts. Replacing an
existing image needs authorization; inspect the reported output rather than
assuming that any file left behind is a complete successful export.

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

Use current help for supported settings rather than copying Excel constants or
an old option catalogue. Replacing a regular chart's source can change all
series, so verify their names, values, and categories afterward.

Choose blank/hidden-cell behavior, axis scales, and any trendline according to
the data's meaning. A different appearance must not imply unsupported
conclusions. Formatting a PivotChart does not authorize changing its fields
or filter scope.

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

The plotted series are not necessarily the same count as the PivotTable's
Values fields: one measure can produce several provider series. Filtering
changes the actual plotted names, categories, and values.

A regular chart's reported source formula is not a complete source rectangle.
Use the plotted data and known source cells for verification rather than
inventing addresses from value arrays. For a PivotChart, verify the actual
linked PivotTable and its fields, filters, and results.

Keep checks and changes within the request. Changing a title or unit display
does not authorize replacing the source, rebuilding unrelated data, restyling
the workbook, adding charts, or repairing unrelated pre-existing errors.
