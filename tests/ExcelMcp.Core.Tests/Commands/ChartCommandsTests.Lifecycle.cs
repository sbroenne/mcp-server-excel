using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Chart lifecycle operations (List, Read, CreateFromRange, CreateFromPivotTable, Delete, Move).
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "true")]
public partial class ChartCommandsTests : IClassFixture<ChartTestsFixture>
{
    private readonly ChartCommands _commands;
    private readonly ChartTestsFixture _fixture;

    public ChartCommandsTests(ChartTestsFixture fixture)
    {
        _commands = new ChartCommands();
        _fixture = fixture;
    }

    [Fact]
    public void List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Act
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);
        var charts = _commands.List(batch);

        // Assert
        Assert.Empty(charts);
    }

    [Fact]
    public void CreateFromRange_ValidData_CreatesChart()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        // Act
        var createResult = _commands.CreateFromRange(
            batch,
            "Sheet1",
            "A1:B4",
            ChartType.ColumnClustered,
            100,
            50,
            400,
            300,
            "TestChart");

        // Assert
        Assert.Equal("TestChart", createResult.ChartName);
        Assert.Equal("Sheet1", createResult.SheetName);
        Assert.Equal(ChartType.ColumnClustered, createResult.ChartType);
        Assert.False(createResult.IsPivotChart);

        // Verify chart exists
        var charts = _commands.List(batch);
        Assert.Single(charts);
        Assert.Equal("TestChart", charts[0].Name);
    }

    [Fact]
    public void CreateFromTable_ValidTable_CreatesChart()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        // Create test data and table
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? listObjects = null;
            dynamic? table = null;

            try
            {
                sheet = ctx.Book.Worksheets[1];

                // Set up data
                sheet.Range["A1:B4"].Value2 = new object[,] {
                    { "Category", "Values" },
                    { "Q1", 100 },
                    { "Q2", 150 },
                    { "Q3", 200 }
                };

                // Create table from range
                listObjects = sheet.ListObjects;
                table = listObjects.Add(1, sheet.Range["A1:B4"], null, 1); // xlYes = 1
                table.Name = "SalesTable";

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref listObjects);
                ComUtilities.Release(ref sheet);
            }
        });

        // Act
        var createResult = _commands.CreateFromTable(
            batch,
            "SalesTable",
            "Sheet1",
            ChartType.ColumnClustered,
            100,
            100,
            400,
            300,
            "TableChart");

        // Assert
        Assert.True(createResult.Success, $"CreateFromTable failed: {createResult.ChartName}");
        Assert.Equal("TableChart", createResult.ChartName);
        Assert.Equal("Sheet1", createResult.SheetName);
        Assert.Equal(ChartType.ColumnClustered, createResult.ChartType);
        Assert.False(createResult.IsPivotChart);

        // Verify chart exists
        var charts = _commands.List(batch);
        Assert.Single(charts);
        Assert.Equal("TableChart", charts[0].Name);
    }

    [Fact]
    public void CreateFromTable_NonExistentTable_ThrowsException()
    {
        // Act & Assert
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _commands.CreateFromTable(
                batch,
                "NonExistentTable",
                "Sheet1",
                ChartType.ColumnClustered,
                50,
                50));

        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CreateFromPivotTable_RangePivotTable_CreatesPivotChart()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        // Create data and PivotTable
        string pivotTableName = "TestPivot";
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? dataRange = null;
            dynamic? pivotCache = null;
            dynamic? newSheet = null;
            dynamic? pivot = null;
            dynamic? rowField = null;
            dynamic? dataField = null;

            try
            {
                sheet = ctx.Book.Worksheets.Item[1];

                // Create sample data
                sheet.Range["A1:C5"].Value2 = new object[,] {
                    { "Product", "Region", "Sales" },
                    { "Widget", "North", 100 },
                    { "Widget", "South", 150 },
                    { "Gadget", "North", 200 },
                    { "Gadget", "South", 250 }
                };

                // Create PivotTable
                dataRange = sheet.Range["A1:C5"];
                pivotCache = ctx.Book.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, dataRange);
                newSheet = ctx.Book.Worksheets.Add();
                newSheet.Name = "PivotSheet";
                pivot = pivotCache.CreatePivotTable(newSheet.Range["A1"], pivotTableName);

                // Add fields
                rowField = pivot.PivotFields("Product");
                rowField.Orientation = 1; // xlRowField

                dataField = pivot.PivotFields("Sales");
                dataField.Orientation = 4; // xlDataField

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref dataField);
                ComUtilities.Release(ref rowField);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref newSheet);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref dataRange);
                ComUtilities.Release(ref sheet);
            }
        });

        // Act
        var result = _commands.CreateFromPivotTable(
            batch,
            pivotTableName,
            "PivotSheet",
            ChartType.ColumnClustered,
            300,
            50,
            400,
            300,
            "PivotChart1");

        // Assert
        Assert.True(result.IsPivotChart, "Chart should be marked as PivotChart");
        Assert.Equal(pivotTableName, result.LinkedPivotTable);
        Assert.Equal("PivotSheet", result.SheetName);
        Assert.Equal(ChartType.ColumnClustered, result.ChartType);

        // Verify chart exists in list
        var charts = _commands.List(batch);
        Assert.Contains(charts, c => c.Name == result.ChartName && c.IsPivotChart);
    }

    [Fact]
    public void CreateFromPivotTable_NonExistentPivotTable_ThrowsException()
    {
        // Act & Assert
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _commands.CreateFromPivotTable(
                batch,
                "NonExistentPivot",
                "Sheet1",
                ChartType.ColumnClustered,
                50,
                50));

        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CreateFromPivotTable_DifferentChartTypes_CreatesCorrectType()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        // Create data and PivotTable
        string pivotTableName = "ChartTypePivot";
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? dataRange = null;
            dynamic? pivotCache = null;
            dynamic? newSheet = null;
            dynamic? pivot = null;
            dynamic? rowField = null;
            dynamic? dataField = null;

            try
            {
                sheet = ctx.Book.Worksheets.Item[1];

                // Create sample data
                sheet.Range["A1:B4"].Value2 = new object[,] {
                    { "Category", "Value" },
                    { "A", 10 },
                    { "B", 20 },
                    { "C", 30 }
                };

                // Create PivotTable
                dataRange = sheet.Range["A1:B4"];
                pivotCache = ctx.Book.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, dataRange);
                newSheet = ctx.Book.Worksheets.Add();
                newSheet.Name = "PivotSheet2";
                pivot = pivotCache.CreatePivotTable(newSheet.Range["A1"], pivotTableName);

                // Add fields
                rowField = pivot.PivotFields("Category");
                rowField.Orientation = 1; // xlRowField

                dataField = pivot.PivotFields("Value");
                dataField.Orientation = 4; // xlDataField

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref dataField);
                ComUtilities.Release(ref rowField);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref newSheet);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref dataRange);
                ComUtilities.Release(ref sheet);
            }
        });

        // Act - Create Pie chart
        var result = _commands.CreateFromPivotTable(
            batch,
            pivotTableName,
            "PivotSheet2",
            ChartType.Pie,
            300,
            50,
            300,
            300);

        // Assert
        Assert.Equal(ChartType.Pie, result.ChartType);
        Assert.True(result.IsPivotChart);
    }

    [Fact]
    public void Read_ExistingChart_ReturnsDetails()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Pie, 50, 50, 300, 300, "PieChart");

        // Act
        var readResult = _commands.Read(batch, "PieChart");

        // Assert
        Assert.Equal("PieChart", readResult.Name);
        Assert.Equal("Sheet1", readResult.SheetName);
        Assert.Equal(ChartType.Pie, readResult.ChartType);
        Assert.False(readResult.IsPivotChart);
        Assert.True(readResult.Series.Count > 0);
    }

    [Fact]
    public void Read_NonExistentChart_ReturnsError()
    {
        // Act & Assert
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);
        var exception = Assert.Throws<InvalidOperationException>(() => _commands.Read(batch, "NonExistent"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Delete_ExistingChart_RemovesChart()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        _commands.CreateFromRange(batch, "Sheet1", "A1:B3", ChartType.Line, 50, 50);

        var chartsBefore = _commands.List(batch);
        Assert.Single(chartsBefore);
        string chartName = chartsBefore[0].Name;

        // Act
        _commands.Delete(batch, chartName);

        // Assert - Verify chart removed
        var chartsAfter = _commands.List(batch);
        Assert.Empty(chartsAfter);
    }

    [Fact]
    public void Move_ExistingChart_UpdatesPosition()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B3", ChartType.ColumnClustered, 100, 100, 300, 200);

        // Act
        _commands.Move(batch, createResult.ChartName, left: 200, top: 150, width: 400, height: 250);

        // Assert - Verify position updated
        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.Equal(200, readResult.Left);
        Assert.Equal(150, readResult.Top);
        Assert.Equal(400, readResult.Width);
        Assert.Equal(250, readResult.Height);
    }

    [Fact]
    public void List_MultipleCharts_ReturnsAll()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        // Create multiple charts
        _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50, 300, 200, "Chart1");
        _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Pie, 400, 50, 300, 200, "Chart2");
        _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Line, 50, 300, 300, 200, "Chart3");

        // Act
        var charts = _commands.List(batch);

        // Assert
        Assert.Equal(3, charts.Count);
        Assert.Contains(charts, c => c.Name == "Chart1" && c.ChartType == ChartType.ColumnClustered);
        Assert.Contains(charts, c => c.Name == "Chart2" && c.ChartType == ChartType.Pie);
        Assert.Contains(charts, c => c.Name == "Chart3" && c.ChartType == ChartType.Line);
    }

    [Fact]
    public void CreateFromRange_DifferentChartTypes_CreatesCorrectly()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        // Act & Assert - Test various chart types
        var chartTypes = new[]
        {
            ChartType.ColumnClustered,
            ChartType.BarClustered,
            ChartType.Line,
            ChartType.Pie,
            ChartType.XYScatter,
            ChartType.Area
        };

        int x = 50;
        foreach (var chartType in chartTypes)
        {
            var result = _commands.CreateFromRange(batch, "Sheet1", "A1:C5", chartType, x, 50, 250, 200);
            Assert.Equal(chartType, result.ChartType);
            x += 300;
        }

        // Verify all created
        var charts = _commands.List(batch);
        Assert.Equal(chartTypes.Length, charts.Count);
    }
}




