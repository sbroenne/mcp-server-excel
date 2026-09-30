# Charts

Chart lifecycle operations create, list, read, move, fit, and delete charts.
Chart-configuration operations manage series, titles, axes, labels, styles, and
trendlines. Reuse the returned chart name rather than assuming Excel's default.

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

## Configuration

- Series indices are 1-based. Adding a series requires a values range; supply its
  category range when the axis labels are not implicit.
- Replacing the source range can change all series. Verify names, values, and
  categories afterward.
- Set per-series chart types for regular combo charts. Use plot options for
  row/column orientation, blanks, and whether hidden cells are plotted.
- Axis titles distinguish Category, Value, and secondary axes. Use US number
  formats for currency/percentage labels.
- Built-in chart styles are 1-48. Area formatting controls chart/plot backgrounds;
  series formatting controls fills, lines, and markers.
- Placement 1 moves and sizes with cells, 2 moves only, 3 is free floating.
- Trendlines include Linear, Exponential, Logarithmic, Polynomial, Power, and
  MovingAverage. Polynomial order is 2-6; moving-average period is at least 2.
  Respect Excel's data/domain requirements for the selected fit.

For multiple charts, use explicit non-overlapping cell ranges with consistent
sizes and spacing. Auto-placement is suitable for a vertical stack. Read the
saved chart's actual geometry and series; a successful creation or a prose
description alone does not establish a correct chart.
