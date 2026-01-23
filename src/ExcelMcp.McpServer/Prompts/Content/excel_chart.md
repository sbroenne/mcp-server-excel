````markdown
# excel_chart - Server Quirks

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

**Related tools**:

| Goal | Tool |
|------|------|
| Configure series, titles, legends | excel_chart_config |
| Create source PivotTable | excel_pivottable |
| Display data as table instead | excel_table create-from-dax |

**Common mistakes**:

- Creating PivotTable + Chart instead of PivotChart (unnecessary complexity)
- Creating chart from range when Data Model data is available
- Forgetting to create PivotTable first for PivotChart

````
