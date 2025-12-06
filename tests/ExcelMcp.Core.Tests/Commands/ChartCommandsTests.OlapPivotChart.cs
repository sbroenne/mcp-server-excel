// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for creating charts from OLAP (Data Model) PivotTables.
/// These tests use <see cref="DataModelPivotTableFixture"/> which creates a workbook
/// with Power Pivot Data Model, DAX measures, and OLAP-based PivotTables.
///
/// This tests the OLAP-specific chart creation path which uses Shapes.AddChart() + SetSourceData()
/// instead of PivotCache.CreatePivotChart() (which fails for OLAP sources).
/// </summary>
[Collection("DataModel")]
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "true")]
public class ChartCommandsOlapTests
{
    private readonly ChartCommands _commands;
    private readonly DataModelPivotTableFixture _fixture;

    public ChartCommandsOlapTests(DataModelPivotTableFixture fixture)
    {
        _commands = new ChartCommands();
        _fixture = fixture;
    }

    [Fact]
    public void CreateFromPivotTable_OlapDataModelPivot_CreatesPivotChart()
    {
        // Arrange - Use the Data Model PivotTable from fixture
        // The fixture creates "DataModelPivot" PivotTable on sheet "ModelData"
        string pivotTableName = "DataModelPivot";
        string sheetName = "ModelData";

        // Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var result = _commands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.ColumnClustered,
            300,
            200,
            400,
            300,
            "OlapChart1");

        // Assert
        Assert.True(result.IsPivotChart, "Chart should be marked as PivotChart");
        Assert.Equal(pivotTableName, result.LinkedPivotTable);
        Assert.Equal(sheetName, result.SheetName);
        Assert.Equal(ChartType.ColumnClustered, result.ChartType);
        Assert.NotNull(result.ChartName);

        // Verify chart exists in list
        var charts = _commands.List(batch);
        Assert.Contains(charts, c => c.Name == result.ChartName);
    }

    [Fact]
    public void CreateFromPivotTable_OlapPivotWithDaxMeasures_CreatesPivotChart()
    {
        // Arrange - Use the Data Model PivotTable that includes DAX measures
        // DataModelPivot has measures like "Total Sales", "Total Revenue", etc.
        string pivotTableName = "DataModelPivot";
        string sheetName = "ModelData";

        // Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var result = _commands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.Pie,
            300,
            400,
            350,
            350,
            "OlapPieChart");

        // Assert
        Assert.True(result.IsPivotChart);
        Assert.Equal(ChartType.Pie, result.ChartType);

        // Verify chart was created
        var chartInfo = _commands.Read(batch, result.ChartName);
        Assert.Equal("OlapPieChart", chartInfo.Name);
        Assert.Equal(sheetName, chartInfo.SheetName);
    }

    [Fact]
    public void CreateFromPivotTable_OlapPivot_BarChart_CreatesCorrectType()
    {
        // Arrange
        string pivotTableName = "DataModelPivot";
        string sheetName = "ModelData";

        // Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var result = _commands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.BarClustered,
            50,
            500,
            400,
            300,
            "OlapBarChart");

        // Assert
        Assert.True(result.IsPivotChart);
        Assert.Equal(ChartType.BarClustered, result.ChartType);

        // Verify via Read
        var chartInfo = _commands.Read(batch, result.ChartName);
        Assert.Equal(ChartType.BarClustered, chartInfo.ChartType);
    }

    [Fact]
    public void CreateFromPivotTable_OlapPivot_LineChart_CreatesCorrectType()
    {
        // Arrange
        string pivotTableName = "DataModelPivot";
        string sheetName = "ModelData";

        // Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var result = _commands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.Line,
            50,
            700,
            400,
            300,
            "OlapLineChart");

        // Assert
        Assert.True(result.IsPivotChart);
        Assert.Equal(ChartType.Line, result.ChartType);

        // Verify chart appears in list
        var charts = _commands.List(batch);
        Assert.Contains(charts, c => c.Name == "OlapLineChart" && c.ChartType == ChartType.Line);
    }

    [Fact]
    public void CreateFromPivotTable_DisambiguationTestPivot_CreatesPivotChart()
    {
        // Arrange - Use the second OLAP PivotTable created by fixture
        // "DisambiguationTest" is on sheet "DisambiguationPivot"
        string pivotTableName = "DisambiguationTest";
        string sheetName = "DisambiguationPivot";

        // Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var result = _commands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.ColumnClustered,
            50,
            200,
            400,
            300,
            "DisambiguationChart");

        // Assert
        Assert.True(result.IsPivotChart);
        Assert.Equal(pivotTableName, result.LinkedPivotTable);
        Assert.Equal(sheetName, result.SheetName);

        // Verify chart exists
        var charts = _commands.List(batch);
        Assert.Contains(charts, c => c.Name == "DisambiguationChart");
    }

    [Fact]
    public void CreateFromPivotTable_OlapPivot_CustomPositionAndSize_AppliesCorrectly()
    {
        // Arrange
        string pivotTableName = "DataModelPivot";
        string sheetName = "ModelData";

        double expectedLeft = 150;
        double expectedTop = 100;
        double expectedWidth = 500;
        double expectedHeight = 400;

        // Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var result = _commands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.ColumnClustered,
            expectedLeft,
            expectedTop,
            expectedWidth,
            expectedHeight,
            "PositionedOlapChart");

        // Assert
        Assert.Equal(expectedLeft, result.Left);
        Assert.Equal(expectedTop, result.Top);
        Assert.Equal(expectedWidth, result.Width);
        Assert.Equal(expectedHeight, result.Height);

        // Verify via Read
        var chartInfo = _commands.Read(batch, result.ChartName);
        Assert.Equal(expectedLeft, chartInfo.Left);
        Assert.Equal(expectedTop, chartInfo.Top);
        Assert.Equal(expectedWidth, chartInfo.Width);
        Assert.Equal(expectedHeight, chartInfo.Height);
    }
}
