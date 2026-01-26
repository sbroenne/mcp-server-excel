````markdown
# excel_chart - Server Quirks

**CRITICAL: Always position charts to avoid overlapping data!**

Before creating a chart:
1. **Check used range**: `excel_range(action: 'get-used-range')` → know where data ends
2. **Use targetRange (PREFERRED)**: Position chart by cell reference in one step
3. **Or use coordinates**: Set `left` and `top` explicitly (don't rely on defaults)

**Positioning with targetRange (recommended)**:
```
excel_chart(create-from-range, sourceRange='A1:B10', chartType='Line', targetRange='F2:K15')
```
Creates chart AND positions it to F2:K15 in one call.

**Positioning with coordinates**:
```
Data in A1:D10 → Chart at row 12 (below) or column F (right)
- Below: top = (lastRow + 2) * 15 points (estimate row height)
- Right: left = (lastCol + 2) * 60 points (estimate col width)
```

---

**Charting Data Model data - use PivotChart directly!**

When visualizing Data Model data:

- **Use**: `excel_chart create-from-pivottable` (creates a PivotChart)
- **NOT**: Create PivotTable → Create separate Chart from PivotTable range

A PivotChart is a single object connected to the Data Model. Creating a PivotTable + separate chart creates two objects unnecessarily.

**Action disambiguation**:

- list: List all charts in workbook (name, type, sheet, data source)
- read: Get chart details (type, position, series, linked PivotTable)
- create-from-range: Create chart from worksheet data range
- create-from-pivottable: Create PivotChart linked to a PivotTable (PREFERRED for Data Model)
- move: Reposition/resize chart
- delete: Remove chart

**PivotChart workflow**:

```
Step 1: Have data in Data Model (via excel_table add-to-datamodel or excel_powerquery)
Step 2: Create PivotTable: excel_pivottable(action: 'create', ...)
Step 3: Create PivotChart: excel_chart(action: 'create-from-pivottable', pivotTableName: '...')
```

Or if you only need the chart (no PivotTable visible):
- Create a PivotTable on a hidden sheet, then create PivotChart from it

**Chart types** (70+ available):

| Category | Common Types |
|----------|--------------|
| Column | ColumnClustered, ColumnStacked, Column3D |
| Line | Line, LineMarkers, LineStacked |
| Pie | Pie, Pie3D, PieExploded, Doughnut |
| Bar | BarClustered, BarStacked |
| Area | Area, AreaStacked |
| Scatter | XYScatter, XYScatterLines |

**Positioning**:

- Units: Points (72 points = 1 inch)
- Default size: 400×300 points
- Use left/top parameters for placement
- **Always check used range first to avoid overlaps**

**Related tools**:

| Goal | Tool |
|------|------|
| Configure series, titles, legends | excel_chart_config |
| Create source PivotTable | excel_pivottable |
| Display data as table instead | excel_table create-from-dax |
| Position chart in empty area | excel_chart_config set-placement or fit-to-range |

**Common mistakes**:

- Creating PivotTable + Chart instead of PivotChart (unnecessary complexity)
- Creating chart from range when Data Model data is available
- Forgetting to create PivotTable first for PivotChart
- **Placing charts at default position (overlaps data!)**

````
