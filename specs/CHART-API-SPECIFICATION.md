# Excel Chart API Specification

> **Comprehensive specification for Excel chart operations - creating and managing Regular Charts and PivotCharts**

## Implementation Status

**Phase 1 (MVP): 🚧 IN PROGRESS** (As of November 22, 2025)
- 🚧 Core interface and strategy pattern design
- 🚧 Regular Chart and PivotChart lifecycle operations
- 🚧 MCP Server integration planning
- ⏸️ CLI commands (after core implementation)
- ⏸️ Integration tests (after core implementation)

**Target: 20-25 core operations covering 90% of chart automation use cases**

---

## Executive Summary

This specification defines a **Chart API** for ExcelMcp that provides complete chart lifecycle management, data source configuration, and appearance customization through Excel COM automation. The API handles **two fundamentally different chart types**:

1. **Regular Charts** - Static charts created from Excel ranges/tables
2. **PivotCharts** - Dynamic interactive charts linked to PivotTables

### Key Design Decisions

1. **COM-Backed Only** - Every operation uses native Excel COM Chart objects
2. **Two-Type Strategy Pattern** - Regular vs PivotChart behavior differences handled transparently
3. **Unified API** - Same method signatures for both chart types (strategy handles implementation)
4. **Complete ChartType Enum** - All 70+ Excel chart types exposed (grouped by category)
5. **Positioning Required** - Both Regular and PivotCharts require left/top/width/height coordinates

### Goals

1. **Complete Lifecycle** - Create, read, move, delete charts of both types
2. **Data Source Management** - Set ranges, add/remove series (Regular), sync with PivotTable (PivotCharts)
3. **Appearance Control** - Chart types, titles, legends, styles
4. **LLM-Friendly** - Clear error messages guide workflow when operations differ between types
5. **90% Coverage** - Core operations satisfy most automation scenarios

---

## Excel Chart Architecture

### What Are Excel Charts?

Excel charts are **visual representations of data** that provide:
- Multiple chart types (column, line, pie, scatter, etc.)
- Embedded positioning on worksheets (left, top, width, height)
- Data series from ranges or PivotTable fields
- Titles, legends, axes, and styling
- Export and presentation capabilities

### The Two Chart Types (Behavioral Difference)

#### Regular Charts
- **Data Source**: Excel ranges or tables
- **Behavior**: Static - updates only when source range updates
- **Series Management**: Explicit via SeriesCollection
- **Creation**: `Shapes.AddChart()` or `ChartObjects.Add()`
- **Use Cases**: Reports, dashboards, static analysis

#### PivotCharts
- **Data Source**: PivotTable/PivotCache
- **Behavior**: Dynamic - updates automatically when PivotTable changes
- **Series Management**: Automatic sync with PivotTable value fields
- **Creation**: `PivotCache.CreatePivotChart()` or `Shapes.AddChart2()` + link
- **Use Cases**: Interactive analysis, drill-down reports, OLAP cubes

### Excel COM Object Model

#### Core Objects
```csharp
// Worksheet-level chart access
dynamic shapes = worksheet.Shapes;
dynamic chartObjects = worksheet.ChartObjects;

// Regular Chart creation (Modern API - Excel 2010+)
dynamic shape = shapes.AddChart(
    XlChartType: 51,  // xlColumnClustered
    Left: 100,
    Top: 100,
    Width: 400,
    Height: 300
);
dynamic chart = shape.Chart;

// PivotChart creation (from PivotTable)
dynamic pivotCache = pivotTable.PivotCache();
dynamic pivotChart = pivotCache.CreatePivotChart(
    Destination: worksheet.Range["H1"]  // Top-left position
);

// Chart object hierarchy
dynamic chart = chartObject.Chart;  // OR shape.Chart
dynamic seriesCollection = chart.SeriesCollection();
dynamic series = seriesCollection.Item(1);
```

#### Chart vs ChartObject
- **ChartObject** - Wrapper providing positioning (Left, Top, Width, Height)
- **Chart** - The actual chart with data, type, titles, legends
- Both Regular and PivotCharts use this dual-object pattern

---

## Proposed Chart API Design

### Design Principles

1. **COM-Backed Only**: Every method uses native Excel COM Chart operations
2. **Strategy Pattern**: `IChartStrategy` with `RegularChartStrategy` and `PivotChartStrategy`
3. **Unified API**: Same method signatures - strategy handles implementation differences
4. **Complete Enum**: All 70+ chart types exposed, grouped by category for readability
5. **Error Guidance**: Clear messages guide LLMs when operations differ between types

### Phase 1: Core Operations (MVP)

```csharp
public interface IChartCommands
{
    // === LIFECYCLE OPERATIONS ===
    
    /// <summary>
    /// Lists all charts in workbook (Regular and PivotCharts)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>List of charts with names, types, sheets, positions, data sources</returns>
    ChartListResult List(IExcelBatch batch);
    
    /// <summary>
    /// Gets complete chart configuration
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart (or shape name)</param>
    /// <returns>Chart type, data source, series info, position, styling</returns>
    ChartInfoResult Read(IExcelBatch batch, string chartName);
    
    /// <summary>
    /// Creates a Regular Chart from an Excel range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name for chart placement</param>
    /// <param name="sourceRange">Source data range (e.g., "A1:D10")</param>
    /// <param name="chartType">Chart type from ChartType enum</param>
    /// <param name="left">Left position in points</param>
    /// <param name="top">Top position in points</param>
    /// <param name="width">Width in points (default: 400)</param>
    /// <param name="height">Height in points (default: 300)</param>
    /// <param name="chartName">Optional name for the chart</param>
    /// <returns>Created chart name and configuration</returns>
    ChartCreateResult CreateFromRange(IExcelBatch batch, 
        string sheetName, string sourceRange, 
        ChartType chartType,
        double left, double top, 
        double width = 400, double height = 300,
        string? chartName = null);
    
    /// <summary>
    /// Creates a PivotChart from an existing PivotTable
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="sheetName">Worksheet name for chart placement</param>
    /// <param name="chartType">Chart type from ChartType enum</param>
    /// <param name="left">Left position in points</param>
    /// <param name="top">Top position in points</param>
    /// <param name="width">Width in points (default: 400)</param>
    /// <param name="height">Height in points (default: 300)</param>
    /// <param name="chartName">Optional name for the chart</param>
    /// <returns>Created PivotChart name and linked PivotTable</returns>
    ChartCreateResult CreateFromPivotTable(IExcelBatch batch,
        string pivotTableName, string sheetName,
        ChartType chartType,
        double left, double top,
        double width = 400, double height = 300,
        string? chartName = null);
    
    /// <summary>
    /// Deletes a chart (Regular or PivotChart)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart to delete</param>
    /// <returns>Operation result</returns>
    OperationResult Delete(IExcelBatch batch, string chartName);
    
    /// <summary>
    /// Moves/resizes a chart
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="left">New left position in points (null = no change)</param>
    /// <param name="top">New top position in points (null = no change)</param>
    /// <param name="width">New width in points (null = no change)</param>
    /// <param name="height">New height in points (null = no change)</param>
    /// <returns>Operation result with new position</returns>
    OperationResult Move(IExcelBatch batch, string chartName,
        double? left = null, double? top = null,
        double? width = null, double? height = null);
    
    // === DATA SOURCE OPERATIONS ===
    
    /// <summary>
    /// Sets data source range for Regular Charts
    /// PivotCharts: Returns error guiding to pivottable
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="sourceRange">New source range (e.g., "Sheet1!A1:D10")</param>
    /// <returns>Operation result</returns>
    OperationResult SetSourceRange(IExcelBatch batch, string chartName, string sourceRange);
    
    /// <summary>
    /// Adds a data series to Regular Charts
    /// PivotCharts: Returns error guiding to pivottable(action: 'add-value-field')
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesName">Name for the series</param>
    /// <param name="valuesRange">Range containing Y values (e.g., "Sheet1!B2:B10")</param>
    /// <param name="categoryRange">Optional range for X values/categories</param>
    /// <returns>Series information</returns>
    ChartSeriesResult AddSeries(IExcelBatch batch, string chartName,
        string seriesName, string valuesRange, string? categoryRange = null);
    
    /// <summary>
    /// Removes a data series from Regular Charts
    /// PivotCharts: Returns error guiding to pivottable(action: 'remove-field')
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based index of series to remove</param>
    /// <returns>Operation result</returns>
    OperationResult RemoveSeries(IExcelBatch batch, string chartName, int seriesIndex);
    
    // === APPEARANCE OPERATIONS ===
    
    /// <summary>
    /// Changes chart type (works for both Regular and PivotCharts)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="chartType">New chart type from ChartType enum</param>
    /// <returns>Operation result</returns>
    OperationResult SetChartType(IExcelBatch batch, string chartName, ChartType chartType);
    
    /// <summary>
    /// Sets chart title
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="title">Chart title text (empty to hide title)</param>
    /// <returns>Operation result</returns>
    OperationResult SetTitle(IExcelBatch batch, string chartName, string title);
    
    /// <summary>
    /// Sets axis title
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Axis type (Primary, Secondary, Category, Value)</param>
    /// <param name="title">Axis title text</param>
    /// <returns>Operation result</returns>
    OperationResult SetAxisTitle(IExcelBatch batch, string chartName, 
        ChartAxisType axis, string title);
    
    /// <summary>
    /// Shows or hides chart legend
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="visible">True to show legend, false to hide</param>
    /// <param name="position">Legend position (optional)</param>
    /// <returns>Operation result</returns>
    OperationResult ShowLegend(IExcelBatch batch, string chartName, 
        bool visible, LegendPosition? position = null);
    
    /// <summary>
    /// Applies a chart style
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="styleId">Style number (1-48)</param>
    /// <returns>Operation result</returns>
    OperationResult SetStyle(IExcelBatch batch, string chartName, int styleId);
}

// === SUPPORTING TYPES ===

/// <summary>
/// Excel chart types - All 70+ values grouped by category
/// Excel COM: XlChartType enumeration
/// Reference: https://learn.microsoft.com/office/vba/api/excel.xlcharttype
/// </summary>
public enum ChartType
{
    // === COLUMN CHARTS ===
    ColumnClustered = 51,           // xlColumnClustered
    ColumnStacked = 52,             // xlColumnStacked
    ColumnStacked100 = 53,          // xlColumnStacked100
    Column3DClustered = 54,         // xl3DColumnClustered
    Column3DStacked = 55,           // xl3DColumnStacked
    Column3DStacked100 = 56,        // xl3DColumnStacked100
    Column3D = -4100,               // xl3DColumn
    
    // === BAR CHARTS ===
    BarClustered = 57,              // xlBarClustered
    BarStacked = 58,                // xlBarStacked
    BarStacked100 = 59,             // xlBarStacked100
    Bar3DClustered = 60,            // xl3DBarClustered
    Bar3DStacked = 61,              // xl3DBarStacked
    Bar3DStacked100 = 62,           // xl3DBarStacked100
    
    // === LINE CHARTS ===
    Line = 4,                       // xlLine
    LineStacked = 63,               // xlLineStacked
    LineStacked100 = 64,            // xlLineStacked100
    LineMarkers = 65,               // xlLineMarkers
    LineMarkersStacked = 66,        // xlLineMarkersStacked
    LineMarkersStacked100 = 67,     // xlLineMarkersStacked100
    Line3D = -4101,                 // xl3DLine
    
    // === PIE CHARTS ===
    Pie = 5,                        // xlPie
    Pie3D = -4102,                  // xl3DPie
    PieOfPie = 68,                  // xlPieOfPie
    PieExploded = 69,               // xlPieExploded
    PieExploded3D = 70,             // xl3DPieExploded
    BarOfPie = 71,                  // xlBarOfPie
    
    // === SCATTER (XY) CHARTS ===
    XYScatter = -4169,              // xlXYScatter
    XYScatterSmooth = 72,           // xlXYScatterSmooth
    XYScatterSmoothNoMarkers = 73,  // xlXYScatterSmoothNoMarkers
    XYScatterLines = 74,            // xlXYScatterLines
    XYScatterLinesNoMarkers = 75,   // xlXYScatterLinesNoMarkers
    
    // === AREA CHARTS ===
    Area = 1,                       // xlArea
    AreaStacked = 76,               // xlAreaStacked
    AreaStacked100 = 77,            // xlAreaStacked100
    Area3D = -4098,                 // xl3DArea
    Area3DStacked = 78,             // xl3DAreaStacked
    Area3DStacked100 = 79,          // xl3DAreaStacked100
    
    // === DOUGHNUT CHARTS ===
    Doughnut = -4120,               // xlDoughnut
    DoughnutExploded = 80,          // xlDoughnutExploded
    
    // === RADAR CHARTS ===
    Radar = -4151,                  // xlRadar
    RadarMarkers = 81,              // xlRadarMarkers
    RadarFilled = 82,               // xlRadarFilled
    
    // === SURFACE CHARTS ===
    Surface = 83,                   // xlSurface
    SurfaceWireframe = 84,          // xlSurfaceWireframe
    SurfaceTopView = 85,            // xlSurfaceTopView
    SurfaceTopViewWireframe = 86,   // xlSurfaceTopViewWireframe
    
    // === BUBBLE CHARTS ===
    Bubble = 15,                    // xlBubble
    Bubble3DEffect = 87,            // xlBubble3DEffect
    
    // === STOCK CHARTS ===
    StockHLC = 88,                  // xlStockHLC (High-Low-Close)
    StockOHLC = 89,                 // xlStockOHLC (Open-High-Low-Close)
    StockVHLC = 90,                 // xlStockVHLC (Volume-High-Low-Close)
    StockVOHLC = 91,                // xlStockVOHLC (Volume-Open-High-Low-Close)
    
    // === CYLINDER CHARTS ===
    CylinderBarClustered = 95,      // xlCylinderBarClustered
    CylinderBarStacked = 96,        // xlCylinderBarStacked
    CylinderBarStacked100 = 97,     // xlCylinderBarStacked100
    CylinderCol = 98,               // xlCylinderCol
    CylinderColClustered = 92,      // xlCylinderColClustered
    CylinderColStacked = 93,        // xlCylinderColStacked
    CylinderColStacked100 = 94,     // xlCylinderColStacked100
    
    // === CONE CHARTS ===
    ConeBarClustered = 102,         // xlConeBarClustered
    ConeBarStacked = 103,           // xlConeBarStacked
    ConeBarStacked100 = 104,        // xlConeBarStacked100
    ConeCol = 105,                  // xlConeCol
    ConeColClustered = 99,          // xlConeColClustered
    ConeColStacked = 100,           // xlConeColStacked
    ConeColStacked100 = 101,        // xlConeColStacked100
    
    // === PYRAMID CHARTS ===
    PyramidBarClustered = 109,      // xlPyramidBarClustered
    PyramidBarStacked = 110,        // xlPyramidBarStacked
    PyramidBarStacked100 = 111,     // xlPyramidBarStacked100
    PyramidCol = 112,               // xlPyramidCol
    PyramidColClustered = 106,      // xlPyramidColClustered
    PyramidColStacked = 107,        // xlPyramidColStacked
    PyramidColStacked100 = 108,     // xlPyramidColStacked100
    
    // === MODERN CHARTS (Excel 2016+) ===
    Treemap = 117,                  // xlTreemap
    Sunburst = 116,                 // xlSunburst
    Histogram = 118,                // xlHistogram
    Pareto = 122,                   // xlPareto
    BoxWhisker = 121,               // xlBoxWhisker
    Waterfall = 119,                // xlWaterfall
    Funnel = 123,                   // xlFunnel
    
    // === COMBO CHARTS ===
    ColumnLineCombo = 120,          // xlColumnLineCombo (approximation)
    RegionMap = 140                 // xlRegionMap (Excel 365)
}

public enum ChartAxisType
{
    Primary,
    Secondary,
    Category,
    Value
}

public enum LegendPosition
{
    Bottom = -4107,     // xlLegendPositionBottom
    Corner = 2,         // xlLegendPositionCorner
    Custom = -4161,     // xlLegendPositionCustom
    Left = -4131,       // xlLegendPositionLeft
    Right = -4152,      // xlLegendPositionRight
    Top = -4160         // xlLegendPositionTop
}

public class ChartListResult
{
    public List<ChartInfo> Charts { get; set; } = new();
    public string FilePath { get; set; } = string.Empty;
}

public class ChartInfo
{
    public string Name { get; set; } = string.Empty;
    public string SheetName { get; set; } = string.Empty;
    public ChartType ChartType { get; set; }
    public bool IsPivotChart { get; set; }
    public string? LinkedPivotTable { get; set; }
    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public int SeriesCount { get; set; }
}

public class ChartInfo
{
    public string Name { get; set; } = string.Empty;
    public string SheetName { get; set; } = string.Empty;
    public ChartType ChartType { get; set; }
    public bool IsPivotChart { get; set; }
    public string? LinkedPivotTable { get; set; }
    public string? SourceRange { get; set; }
    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public string? Title { get; set; }
    public bool HasLegend { get; set; }
    public List<SeriesInfo> Series { get; set; } = new();
}

public class SeriesInfo
{
    public string Name { get; set; } = string.Empty;
    public string ValuesRange { get; set; } = string.Empty;
    public string? CategoryRange { get; set; }
}

public class ChartCreateResult : OperationResult
{
    public string ChartName { get; set; } = string.Empty;
    public string SheetName { get; set; } = string.Empty;
    public ChartType ChartType { get; set; }
    public bool IsPivotChart { get; set; }
    public string? LinkedPivotTable { get; set; }
    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
}

public class ChartSeriesResult : OperationResult
{
    public string SeriesName { get; set; } = string.Empty;
    public string ValuesRange { get; set; } = string.Empty;
    public string? CategoryRange { get; set; }
    public int SeriesIndex { get; set; }
}
```

---

## Strategy Pattern Design

### IChartStrategy Interface

```csharp
public interface IChartStrategy
{
    /// <summary>
    /// Determines if this strategy can handle the chart
    /// </summary>
    bool CanHandle(dynamic chart);
    
    /// <summary>
    /// Gets chart information
    /// </summary>
    ChartInfo GetInfo(dynamic chart, string chartName);
    
    /// <summary>
    /// Sets the data source (Regular: range, PivotChart: error)
    /// </summary>
    OperationResult SetSourceRange(dynamic chart, string sourceRange);
    
    /// <summary>
    /// Adds a series (Regular: SeriesCollection, PivotChart: error)
    /// </summary>
    ChartSeriesResult AddSeries(dynamic chart, string seriesName, 
        string valuesRange, string? categoryRange);
    
    /// <summary>
    /// Removes a series (Regular: SeriesCollection, PivotChart: error)
    /// </summary>
    OperationResult RemoveSeries(dynamic chart, int seriesIndex);
}
```

### RegularChartStrategy

```csharp
public class RegularChartStrategy : IChartStrategy
{
    public bool CanHandle(dynamic chart)
    {
        // Regular charts: chart.PivotLayout is null or doesn't exist
        try
        {
            var pivotLayout = chart.PivotLayout;
            return pivotLayout == null;
        }
        catch
        {
            return true; // No PivotLayout = Regular chart
        }
    }
    
    public OperationResult SetSourceRange(dynamic chart, string sourceRange)
    {
        // Use chart.SetSourceData(range, plotBy)
        chart.SetSourceData(sourceRange);
        return new OperationResult { Success = true };
    }
    
    public ChartSeriesResult AddSeries(dynamic chart, string seriesName, 
        string valuesRange, string? categoryRange)
    {
        dynamic seriesCollection = chart.SeriesCollection();
        dynamic newSeries = seriesCollection.NewSeries();
        newSeries.Name = seriesName;
        newSeries.Values = valuesRange;
        if (categoryRange != null)
        {
            newSeries.XValues = categoryRange;
        }
        // ... return result
    }
}
```

### PivotChartStrategy

```csharp
public class PivotChartStrategy : IChartStrategy
{
    public bool CanHandle(dynamic chart)
    {
        // PivotCharts: chart.PivotLayout exists
        try
        {
            var pivotLayout = chart.PivotLayout;
            return pivotLayout != null;
        }
        catch
        {
            return false;
        }
    }
    
    public OperationResult SetSourceRange(dynamic chart, string sourceRange)
    {
        // PivotCharts can't change source - return helpful error
        return new OperationResult
        {
            Success = false,
            ErrorMessage = "Cannot set source range for PivotChart. " +
                          "PivotCharts automatically sync with their PivotTable data source. " +
                          "To modify data, use pivottable tool to update the linked PivotTable."
        };
    }
    
    public ChartSeriesResult AddSeries(dynamic chart, string seriesName, 
        string valuesRange, string? categoryRange)
    {
        // PivotCharts auto-sync with PivotTable fields - return helpful error
        string pivotTableName = chart.PivotLayout.PivotTable.Name;
        return new ChartSeriesResult
        {
            Success = false,
            ErrorMessage = $"Cannot add series directly to PivotChart. " +
                          $"PivotCharts automatically sync with PivotTable '{pivotTableName}' fields. " +
                          $"Use pivottable(action: 'add-value-field', pivotTableName: '{pivotTableName}', fieldName: '<field>') " +
                          $"to add data series."
        };
    }
}
```

---

## Excel COM Implementation Details

### Chart Creation Patterns

#### Regular Chart (Modern API - Excel 2010+)

```csharp
// Using Shapes.AddChart (recommended)
dynamic shapes = worksheet.Shapes;
dynamic shape = shapes.AddChart(
    XlChartType: 51,  // xlColumnClustered
    Left: 100,
    Top: 100,
    Width: 400,
    Height: 300
);
dynamic chart = shape.Chart;

// Set data source
chart.SetSourceData(worksheet.Range["A1:D10"]);

// Optional: Name the chart
shape.Name = "SalesChart";
```

#### PivotChart Creation

```csharp
// Method 1: From PivotCache
dynamic pivotTable = FindPivotTable(workbook, "SalesPivot");
dynamic pivotCache = pivotTable.PivotCache();

// CreatePivotChart returns Shape object containing the chart
dynamic pivotChartShape = pivotCache.CreatePivotChart(
    Destination: worksheet.Range["H1"]  // Top-left corner
);

// Access the chart
dynamic pivotChart = pivotChartShape.Chart;

// Set chart type
pivotChart.ChartType = 51;  // xlColumnClustered

// Position/resize
pivotChartShape.Left = 500;
pivotChartShape.Top = 100;
pivotChartShape.Width = 400;
pivotChartShape.Height = 300;
```

### Chart Detection (Regular vs PivotChart)

```csharp
public static bool IsPivotChart(dynamic chart)
{
    try
    {
        var pivotLayout = chart.PivotLayout;
        return pivotLayout != null;
    }
    catch
    {
        return false; // No PivotLayout property = Regular chart
    }
}
```

### Finding Charts

```csharp
// Find by shape name
dynamic shapes = worksheet.Shapes;
for (int i = 1; i <= shapes.Count; i++)
{
    dynamic shape = shapes.Item(i);
    if (shape.Type == 3)  // msoChart = 3
    {
        if (shape.Name == chartName)
        {
            return shape.Chart;
        }
    }
}

// OR use ChartObjects collection (legacy but still works)
dynamic chartObjects = worksheet.ChartObjects;
for (int i = 1; i <= chartObjects.Count; i++)
{
    dynamic chartObject = chartObjects.Item(i);
    if (chartObject.Name == chartName)
    {
        return chartObject.Chart;
    }
}
```

---

## MCP Tool: chart

### Actions (20-25 operations)

```typescript
{
  "name": "chart",
  "description": "Excel chart operations - create and manage Regular Charts and PivotCharts",
  "parameters": {
    "action": "enum<ChartAction>",
    "excelPath": "string",
    "sessionId": "string",
    "chartName": "string",
    "sheetName": "string",
    "sourceRange": "string",
    "pivotTableName": "string",
    "chartType": "enum<ChartType>",
    "left": "double",
    "top": "double",
    "width": "double",
    "height": "double",
    "seriesName": "string",
    "valuesRange": "string",
    "categoryRange": "string",
    "seriesIndex": "int",
    "title": "string",
    "axis": "enum<ChartAxisType>",
    "visible": "boolean",
    "legendPosition": "enum<LegendPosition>",
    "styleId": "int"
  },
  "actions": [
    // Lifecycle (7 ops)
    "list",                     // List all charts
    "read",                     // Get chart details
    "create-from-range",        // Create Regular Chart
    "create-from-pivottable",   // Create PivotChart
    "delete",                   // Delete chart
    "move",                     // Move/resize chart
    
    // Data Source (3 ops)
    "set-source-range",         // Set data source (Regular only)
    "add-series",               // Add series (Regular only, PivotChart returns error)
    "remove-series",            // Remove series (Regular only, PivotChart returns error)
    
    // Appearance (5 ops)
    "set-chart-type",           // Change chart type
    "set-title",                // Set chart title
    "set-axis-title",           // Set axis title
    "show-legend",              // Show/hide legend
    "set-style"                 // Apply style (1-48)
  ]
}
```

---

## CLI Commands

```powershell
# === LIFECYCLE ===
excelcli chart list <session-id>
excelcli chart read <session-id> <chart-name>
excelcli chart create-from-range <session-id> <sheet> <range> <type> <left> <top> [width] [height] [name]
excelcli chart create-from-pivottable <session-id> <pivot-name> <sheet> <type> <left> <top> [width] [height] [name]
excelcli chart delete <session-id> <chart-name>
excelcli chart move <session-id> <chart-name> [left] [top] [width] [height]

# === DATA SOURCE ===
excelcli chart set-source-range <session-id> <chart-name> <range>
excelcli chart add-series <session-id> <chart-name> <series-name> <values-range> [category-range]
excelcli chart remove-series <session-id> <chart-name> <series-index>

# === APPEARANCE ===
excelcli chart set-chart-type <session-id> <chart-name> <type>
excelcli chart set-title <session-id> <chart-name> <title>
excelcli chart set-axis-title <session-id> <chart-name> <axis> <title>
excelcli chart show-legend <session-id> <chart-name> <true|false> [position]
excelcli chart set-style <session-id> <chart-name> <style-id>
```

---

## Usage Examples

### Creating Regular Chart

```csharp
// Create column chart from range
var result = await chartCommands.CreateFromRange(
    batch, 
    sheetName: "Data",
    sourceRange: "A1:D10",
    chartType: ChartType.ColumnClustered,
    left: 100,
    top: 100,
    width: 400,
    height: 300,
    chartName: "SalesChart"
);

// Customize appearance
await chartCommands.SetTitle(batch, "SalesChart", "Sales by Region");
await chartCommands.SetAxisTitle(batch, "SalesChart", ChartAxisType.Value, "Revenue ($)");
await chartCommands.ShowLegend(batch, "SalesChart", true, LegendPosition.Bottom);
```

### Creating PivotChart

```csharp
// Create PivotChart from existing PivotTable
var result = await chartCommands.CreateFromPivotTable(
    batch,
    pivotTableName: "SalesPivot",
    sheetName: "Dashboard",
    chartType: ChartType.ColumnClustered,
    left: 500,
    top: 100
);

// PivotChart automatically syncs with PivotTable
// To add data series, use pivottable tool:
await pivotCommands.AddValueField(batch, "SalesPivot", "Revenue", AggregationFunction.Sum);
// PivotChart updates automatically!
```

---

## Success Criteria

### Phase 1 (MVP) - 🚧 IN PROGRESS

**Lifecycle Operations (6/6):**
- 🚧 `List` - List all charts
- 🚧 `Read` - Get chart configuration
- 🚧 `CreateFromRange` - Create Regular Chart
- 🚧 `CreateFromPivotTable` - Create PivotChart
- 🚧 `Delete` - Delete chart
- 🚧 `Move` - Move/resize chart

**Data Source Operations (3/3):**
- 🚧 `SetSourceRange` - Set range (Regular only)
- 🚧 `AddSeries` - Add series (Regular only, PivotChart error)
- 🚧 `RemoveSeries` - Remove series (Regular only, PivotChart error)

**Appearance Operations (5/5):**
- 🚧 `SetChartType` - Change chart type (all 70+ types)
- 🚧 `SetTitle` - Set chart title
- 🚧 `SetAxisTitle` - Set axis title
- 🚧 `ShowLegend` - Show/hide legend
- 🚧 `SetStyle` - Apply chart style

**Integration:**
- 🚧 MCP Server tool (`chart` with ~15 actions)
- 🚧 CLI commands (all 15 operations)
- 🚧 Integration tests with both chart types

### Future Enhancements (Phase 2)

- Advanced formatting (colors, fonts, borders)
- Axis scaling and formatting
- Data labels
- Trendlines


---

## Implementation Timeline

**Phase 1 (Core Operations): 🚧 IN PROGRESS** (November 22, 2025 - December 2025)
- Specification and interface design
- Strategy pattern implementation
- Core lifecycle and appearance operations
- MCP Server and CLI integration
- Integration tests
- **Estimated Time:** 1-2 weeks

**Phase 2 (Advanced Features): ⏸️ FUTURE** (On demand)
- Advanced formatting and customization
- Chart export capabilities
- **Estimated Time:** 1 week when prioritized

---

## Open Questions

1. **Chart naming** - Excel auto-generates names like "Chart 1". Should we force users to provide names or auto-generate meaningful ones?

2. **Default positioning** - Should we have a smart default positioning algorithm (e.g., place charts below/beside data) or always require explicit coordinates?

3. **Chart refresh** - PivotCharts auto-refresh. Should we expose a `Refresh` operation for Regular Charts to re-read source data?

4. **Chart export** - Should chart-to-image export be Phase 1 or Phase 2?

**Recommended Answers:**
1. **Optional names** - Auto-generate meaningful names like "Chart_Data_A1D10" if not provided
2. **Explicit coordinates** - Always require positioning (LLMs can calculate, prevents unexpected placement)
3. **No Refresh for Regular** - Charts update automatically when data changes. Not needed.
4. **Phase 2** - Export is advanced feature, not core automation
