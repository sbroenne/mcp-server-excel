# Excel Charts Reference

## Tools

- **`excel_chart`**: Create charts, manage positioning and data sources
- **`excel_chart_config`**: Configure chart appearance, formatting, and analysis features

## Chart Creation

### From Range
```
excel_chart(create-from-range, chartType, sourceRange, sheetName)
```
Best for: Simple data in worksheet ranges

### From PivotTable (PivotChart)
```
excel_chart(create-from-pivottable, pivotTableName)
```
Best for: Data Model data - creates a single PivotChart object (don't create separate PivotTable + Chart)

### From Table
```
excel_chart(create-from-table, tableName, chartType)
```
Best for: Excel Tables with structured references

## Chart Types

Common types: `ColumnClustered`, `Line`, `Pie`, `Bar`, `Area`, `XYScatter`, `Doughnut`

Specialized: `Waterfall`, `Funnel`, `Treemap`, `Sunburst`, `BoxWhisker`, `Histogram`, `Pareto`

## Configuration Actions (excel_chart_config)

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
1. excel_chart(create-from-range) → chartName
2. excel_chart_config(set-title, title="Monthly Sales")
3. excel_chart_config(set-axis-title, axis="Value", title="Revenue ($)")
4. excel_chart_config(set-axis-number-format, axis="Value", numberFormat="$#,##0")
5. excel_chart_config(set-data-labels, position="OutsideEnd", showValue=true)
```

### Add Analysis
```
1. excel_chart_config(add-trendline, trendlineType="Linear", displayEquation=true, displayRSquared=true)
2. excel_chart_config(set-trendline, forward=3) # Forecast 3 periods ahead
```

## Best Practices

1. **PivotCharts for Data Model**: Use `create-from-pivottable` not PivotTable + separate chart
2. **Format numbers**: Set axis number format for readability
3. **Use gridlines sparingly**: Minor gridlines often add clutter
4. **Trendlines for insights**: Add R² to show fit quality
5. **Data labels placement**: `OutsideEnd` for bar charts, `Center` for pie charts
