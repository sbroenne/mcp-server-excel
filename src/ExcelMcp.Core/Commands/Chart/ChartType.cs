namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Excel chart types - All 70+ values grouped by category.
/// Excel COM: XlChartType enumeration.
/// Reference: https://learn.microsoft.com/office/vba/api/excel.xlcharttype
/// </summary>
public enum ChartType
{
    // === COLUMN CHARTS ===

    /// <summary>Clustered column chart (xlColumnClustered)</summary>
    ColumnClustered = 51,

    /// <summary>Stacked column chart (xlColumnStacked)</summary>
    ColumnStacked = 52,

    /// <summary>100% stacked column chart (xlColumnStacked100)</summary>
    ColumnStacked100 = 53,

    /// <summary>3D clustered column chart (xl3DColumnClustered)</summary>
    Column3DClustered = 54,

    /// <summary>3D stacked column chart (xl3DColumnStacked)</summary>
    Column3DStacked = 55,

    /// <summary>3D 100% stacked column chart (xl3DColumnStacked100)</summary>
    Column3DStacked100 = 56,

    /// <summary>3D column chart (xl3DColumn)</summary>
    Column3D = -4100,

    // === BAR CHARTS ===

    /// <summary>Clustered bar chart (xlBarClustered)</summary>
    BarClustered = 57,

    /// <summary>Stacked bar chart (xlBarStacked)</summary>
    BarStacked = 58,

    /// <summary>100% stacked bar chart (xlBarStacked100)</summary>
    BarStacked100 = 59,

    /// <summary>3D clustered bar chart (xl3DBarClustered)</summary>
    Bar3DClustered = 60,

    /// <summary>3D stacked bar chart (xl3DBarStacked)</summary>
    Bar3DStacked = 61,

    /// <summary>3D 100% stacked bar chart (xl3DBarStacked100)</summary>
    Bar3DStacked100 = 62,

    // === LINE CHARTS ===

    /// <summary>Line chart (xlLine)</summary>
    Line = 4,

    /// <summary>Stacked line chart (xlLineStacked)</summary>
    LineStacked = 63,

    /// <summary>100% stacked line chart (xlLineStacked100)</summary>
    LineStacked100 = 64,

    /// <summary>Line chart with markers (xlLineMarkers)</summary>
    LineMarkers = 65,

    /// <summary>Stacked line chart with markers (xlLineMarkersStacked)</summary>
    LineMarkersStacked = 66,

    /// <summary>100% stacked line chart with markers (xlLineMarkersStacked100)</summary>
    LineMarkersStacked100 = 67,

    /// <summary>3D line chart (xl3DLine)</summary>
    Line3D = -4101,

    // === PIE CHARTS ===

    /// <summary>Pie chart (xlPie)</summary>
    Pie = 5,

    /// <summary>3D pie chart (xl3DPie)</summary>
    Pie3D = -4102,

    /// <summary>Pie of pie chart (xlPieOfPie)</summary>
    PieOfPie = 68,

    /// <summary>Exploded pie chart (xlPieExploded)</summary>
    PieExploded = 69,

    /// <summary>3D exploded pie chart (xl3DPieExploded)</summary>
    PieExploded3D = 70,

    /// <summary>Bar of pie chart (xlBarOfPie)</summary>
    BarOfPie = 71,

    // === SCATTER (XY) CHARTS ===

    /// <summary>Scatter chart (xlXYScatter)</summary>
    XYScatter = -4169,

    /// <summary>Scatter chart with smooth lines (xlXYScatterSmooth)</summary>
    XYScatterSmooth = 72,

    /// <summary>Scatter chart with smooth lines and no markers (xlXYScatterSmoothNoMarkers)</summary>
    XYScatterSmoothNoMarkers = 73,

    /// <summary>Scatter chart with lines (xlXYScatterLines)</summary>
    XYScatterLines = 74,

    /// <summary>Scatter chart with lines and no markers (xlXYScatterLinesNoMarkers)</summary>
    XYScatterLinesNoMarkers = 75,

    // === AREA CHARTS ===

    /// <summary>Area chart (xlArea)</summary>
    Area = 1,

    /// <summary>Stacked area chart (xlAreaStacked)</summary>
    AreaStacked = 76,

    /// <summary>100% stacked area chart (xlAreaStacked100)</summary>
    AreaStacked100 = 77,

    /// <summary>3D area chart (xl3DArea)</summary>
    Area3D = -4098,

    /// <summary>3D stacked area chart (xl3DAreaStacked)</summary>
    Area3DStacked = 78,

    /// <summary>3D 100% stacked area chart (xl3DAreaStacked100)</summary>
    Area3DStacked100 = 79,

    // === DOUGHNUT CHARTS ===

    /// <summary>Doughnut chart (xlDoughnut)</summary>
    Doughnut = -4120,

    /// <summary>Exploded doughnut chart (xlDoughnutExploded)</summary>
    DoughnutExploded = 80,

    // === RADAR CHARTS ===

    /// <summary>Radar chart (xlRadar)</summary>
    Radar = -4151,

    /// <summary>Radar chart with markers (xlRadarMarkers)</summary>
    RadarMarkers = 81,

    /// <summary>Filled radar chart (xlRadarFilled)</summary>
    RadarFilled = 82,

    // === SURFACE CHARTS ===

    /// <summary>Surface chart (xlSurface)</summary>
    Surface = 83,

    /// <summary>Wireframe surface chart (xlSurfaceWireframe)</summary>
    SurfaceWireframe = 84,

    /// <summary>Top view surface chart (xlSurfaceTopView)</summary>
    SurfaceTopView = 85,

    /// <summary>Top view wireframe surface chart (xlSurfaceTopViewWireframe)</summary>
    SurfaceTopViewWireframe = 86,

    // === BUBBLE CHARTS ===

    /// <summary>Bubble chart (xlBubble)</summary>
    Bubble = 15,

    /// <summary>Bubble chart with 3D effect (xlBubble3DEffect)</summary>
    Bubble3DEffect = 87,

    // === STOCK CHARTS ===

    /// <summary>Stock chart (High-Low-Close) (xlStockHLC)</summary>
    StockHLC = 88,

    /// <summary>Stock chart (Open-High-Low-Close) (xlStockOHLC)</summary>
    StockOHLC = 89,

    /// <summary>Stock chart (Volume-High-Low-Close) (xlStockVHLC)</summary>
    StockVHLC = 90,

    /// <summary>Stock chart (Volume-Open-High-Low-Close) (xlStockVOHLC)</summary>
    StockVOHLC = 91,

    // === CYLINDER CHARTS ===

    /// <summary>Clustered cylinder bar chart (xlCylinderBarClustered)</summary>
    CylinderBarClustered = 95,

    /// <summary>Stacked cylinder bar chart (xlCylinderBarStacked)</summary>
    CylinderBarStacked = 96,

    /// <summary>100% stacked cylinder bar chart (xlCylinderBarStacked100)</summary>
    CylinderBarStacked100 = 97,

    /// <summary>Cylinder column chart (xlCylinderCol)</summary>
    CylinderCol = 98,

    /// <summary>Clustered cylinder column chart (xlCylinderColClustered)</summary>
    CylinderColClustered = 92,

    /// <summary>Stacked cylinder column chart (xlCylinderColStacked)</summary>
    CylinderColStacked = 93,

    /// <summary>100% stacked cylinder column chart (xlCylinderColStacked100)</summary>
    CylinderColStacked100 = 94,

    // === CONE CHARTS ===

    /// <summary>Clustered cone bar chart (xlConeBarClustered)</summary>
    ConeBarClustered = 102,

    /// <summary>Stacked cone bar chart (xlConeBarStacked)</summary>
    ConeBarStacked = 103,

    /// <summary>100% stacked cone bar chart (xlConeBarStacked100)</summary>
    ConeBarStacked100 = 104,

    /// <summary>Cone column chart (xlConeCol)</summary>
    ConeCol = 105,

    /// <summary>Clustered cone column chart (xlConeColClustered)</summary>
    ConeColClustered = 99,

    /// <summary>Stacked cone column chart (xlConeColStacked)</summary>
    ConeColStacked = 100,

    /// <summary>100% stacked cone column chart (xlConeColStacked100)</summary>
    ConeColStacked100 = 101,

    // === PYRAMID CHARTS ===

    /// <summary>Clustered pyramid bar chart (xlPyramidBarClustered)</summary>
    PyramidBarClustered = 109,

    /// <summary>Stacked pyramid bar chart (xlPyramidBarStacked)</summary>
    PyramidBarStacked = 110,

    /// <summary>100% stacked pyramid bar chart (xlPyramidBarStacked100)</summary>
    PyramidBarStacked100 = 111,

    /// <summary>Pyramid column chart (xlPyramidCol)</summary>
    PyramidCol = 112,

    /// <summary>Clustered pyramid column chart (xlPyramidColClustered)</summary>
    PyramidColClustered = 106,

    /// <summary>Stacked pyramid column chart (xlPyramidColStacked)</summary>
    PyramidColStacked = 107,

    /// <summary>100% stacked pyramid column chart (xlPyramidColStacked100)</summary>
    PyramidColStacked100 = 108,

    // === MODERN CHARTS (Excel 2016+) ===

    /// <summary>Treemap chart (xlTreemap)</summary>
    Treemap = 117,

    /// <summary>Sunburst chart (xlSunburst)</summary>
    Sunburst = 116,

    /// <summary>Histogram chart (xlHistogram)</summary>
    Histogram = 118,

    /// <summary>Pareto chart (xlPareto)</summary>
    Pareto = 122,

    /// <summary>Box and whisker chart (xlBoxWhisker)</summary>
    BoxWhisker = 121,

    /// <summary>Waterfall chart (xlWaterfall)</summary>
    Waterfall = 119,

    /// <summary>Funnel chart (xlFunnel)</summary>
    Funnel = 123,

    // === COMBO CHARTS ===

    /// <summary>Column-line combo chart (xlColumnLineCombo - approximation)</summary>
    ColumnLineCombo = 120,

    /// <summary>Region map chart (xlRegionMap - Excel 365)</summary>
    RegionMap = 140
}


