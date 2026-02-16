# Excel Charts Reference

## Tools

- **`chart`**: Create charts, manage positioning and data sources
- **`chart_config`**: Configure chart appearance, formatting, and analysis features

## Chart Creation

### From Range
```
chart(create-from-range, chartType, sourceRange, sheetName)
```
Best for: Simple data in worksheet ranges

### From PivotTable (PivotChart)
```
chart(create-from-pivottable, pivotTableName)
```
Best for: Data Model data - creates a single PivotChart object (don't create separate PivotTable + Chart)

### From Table
```
chart(create-from-table, tableName, chartType)
```
Best for: Excel Tables with structured references

## Chart Types

Common types: `ColumnClustered`, `Line`, `Pie`, `Bar`, `Area`, `XYScatter`, `Doughnut`

Specialized: `Waterfall`, `Funnel`, `Treemap`, `Sunburst`, `BoxWhisker`, `Histogram`, `Pareto`

## Configuration Actions (chart_config)

### Series Management
- `add-series`: Add data series with valuesRange and optional categoryRange
- `remove-series`: Remove series by index (1-based)
- `set-source-range`: Replace entire chart data source

### Titles and Labels
- `set-title`: Set chart title (empty string hides)
- `set-axis-title`: Set axis labels (Category, Value, CategorySecondary, ValueSecondary)
- `set-data-labels`: Configure data labels (position, showValue, showCategory, showPercentage, showSeriesName, showLegendKey)

### Axis Formatting
- `get-axis-scale`: Get min, max, majorUnit, minorUnit, and auto flags
- `set-axis-scale`: Configure scale properties
- `get-axis-number-format`: Get current tick label format
- `set-axis-number-format`: Format axis numbers (e.g., `"$#,##0,,\"M\""` for millions)

### Gridlines
- `get-gridlines`: Check visibility state
- `set-gridlines`: Show/hide major/minor gridlines

### Series Formatting
- `set-series-format`: Configure markers (style, size, foregroundColor, backgroundColor)

### Trendlines
- `list-trendlines`: View all trendlines on a series
- `add-trendline`: Add Linear, Exponential, Logarithmic, Polynomial, Power, or MovingAverage trendline
- `delete-trendline`: Remove trendline by index
- `set-trendline`: Configure display (equation, R² value) and forecasting (forward, backward periods)

### Styling
- `show-legend`: Control legend visibility and position (Bottom, Corner, Top, Right, Left)
- `set-style`: Apply Excel chart styles (1-48)

## Trendline Details

### Types
| Type | Use Case | Requirements |
|------|----------|--------------|
| Linear | Straight-line trends | None |
| Exponential | Growth/decay patterns | Positive values |
| Logarithmic | Rapid initial change | Positive values |
| Polynomial | Curves with peaks/valleys | Order parameter (2-6) |
| Power | Accelerating rates | Positive values |
| MovingAverage | Smooth fluctuations | Period parameter (2+) |

### Parameters
- **order**: Required for Polynomial (2-6, default 2)
- **period**: Required for MovingAverage (2+, default 2)
- **forward/backward**: Forecast periods ahead/behind data
- **intercept**: Force trend through specific Y value
- **displayEquation**: Show formula on chart
- **displayRSquared**: Show R² goodness-of-fit value

## Common Workflows

### Create Chart with Formatting
```
1. chart(create-from-range) → chartName
2. chart_config(set-title, title="Monthly Sales")
3. chart_config(set-axis-title, axis="Value", title="Revenue ($)")
4. chart_config(set-axis-number-format, axis="Value", numberFormat="$#,##0")
5. chart_config(set-data-labels, position="OutsideEnd", showValue=true)
```

### Add Analysis
```
1. chart_config(add-trendline, trendlineType="Linear", displayEquation=true, displayRSquared=true)
2. chart_config(set-trendline, forward=3) # Forecast 3 periods ahead
```

## Best Practices

1. **PivotCharts for Data Model**: Use `create-from-pivottable` not PivotTable + separate chart
2. **Format numbers**: Set axis number format for readability
3. **Use gridlines sparingly**: Minor gridlines often add clutter
4. **Trendlines for insights**: Add R² to show fit quality
5. **Data labels placement**: `OutsideEnd` for bar charts, `Center` for pie charts

## Chart Positioning (CRITICAL)

**ALWAYS position charts to avoid overlapping data and other charts:**

### Use targetRange (PREFERRED - One Step)
```
chart(create-from-range, sourceRange='A1:B10', chartType='Line', targetRange='F2:K15')
```
Creates chart AND positions it to cell range in one call.

### Check Used Range First
```
range(action: 'get-used-range') → e.g., "Sheet1!A1:D20"
```

### Position with Coordinates
```
chart(create-from-range, sourceRange: 'A1:B10', left: 360, top: 20)
# left/top in points (72 points = 1 inch)
```

### Use FitToRange (After Creation)
```
chart(create-from-range, ...) -> chartName
chart(fit-to-range, chartName, rangeAddress: 'F2:K15')
# Reposition existing chart to cell range
```

### Position Estimates
- Rows: ~15 points per row (varies with row height)
- Columns: ~60 points per column (varies with column width)
- Default chart: 400×300 points

### Positioning Workflow
1. `get-used-range` → Identify data boundaries
2. **Option A (Preferred)**: Use `targetRange='F2:K15'` in create call
3. **Option B**: Calculate `(lastRow + 2) * 15` for top, or `(lastCol + 2) * 60` for left
4. **Verify**: Use `screenshot(capture-sheet)` to visually confirm no overlaps

## Multi-Chart Layout (CRITICAL)

When creating dashboards with multiple charts, **every chart needs explicit positioning**:

### Grid Layout Pattern
```
Data at A1:D10. Place 4 charts in a 2×2 grid below data:

chart(create-from-range, ..., targetRange='A12:F25')   # Top-left
chart(create-from-range, ..., targetRange='G12:L25')   # Top-right
chart(create-from-range, ..., targetRange='A27:F40')   # Bottom-left
chart(create-from-range, ..., targetRange='G27:L40')   # Bottom-right
screenshot(capture-sheet) → Verify no overlaps
```

### Rules
- **Never use default positioning** for 2+ charts — they will stack on top of each other
- Leave at least 1-2 rows/columns gap between charts
- Use `targetRange` on every chart creation call
- If overlap detected in screenshot, use `chart(fit-to-range)` to fix
- Take a final `screenshot(capture-sheet)` to verify the complete layout
