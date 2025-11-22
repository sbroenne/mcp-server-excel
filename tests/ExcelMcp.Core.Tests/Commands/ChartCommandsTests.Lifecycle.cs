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
public partial class ChartCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly ChartCommands _commands;
    private readonly string _tempDir;

    public ChartCommandsTests(TempDirectoryFixture fixture)
    {
        _commands = new ChartCommands();
        _tempDir = fixture.TempDir;
    }

    [Fact]
    public void List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(List_EmptyWorkbook_ReturnsEmptyList),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.Empty(result.Charts);
    }

    [Fact]
    public void CreateFromRange_ValidData_CreatesChart()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(CreateFromRange_ValidData_CreatesChart),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Category";
            sheet.Range["B1"].Value2 = "Values";
            sheet.Range["A2"].Value2 = "Q1";
            sheet.Range["B2"].Value2 = 100;
            sheet.Range["A3"].Value2 = "Q2";
            sheet.Range["B3"].Value2 = 150;
            sheet.Range["A4"].Value2 = "Q3";
            sheet.Range["B4"].Value2 = 200;
            return 0;
        });

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
        Assert.True(createResult.Success, $"CreateFromRange failed: {createResult.ErrorMessage}");
        Assert.Equal("TestChart", createResult.ChartName);
        Assert.Equal("Sheet1", createResult.SheetName);
        Assert.Equal(ChartType.ColumnClustered, createResult.ChartType);
        Assert.False(createResult.IsPivotChart);

        // Verify chart exists
        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);
        Assert.Single(listResult.Charts);
        Assert.Equal("TestChart", listResult.Charts[0].Name);
    }

    [Fact]
    public void Read_ExistingChart_ReturnsDetails()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(Read_ExistingChart_ReturnsDetails),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "Cat", "Val" },
                { "Q1", 100 },
                { "Q2", 150 },
                { "Q3", 200 }
            };
            return 0;
        });

        _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Pie, 50, 50, 300, 300, "PieChart");

        // Act
        var readResult = _commands.Read(batch, "PieChart");

        // Assert
        Assert.True(readResult.Success, $"Read failed: {readResult.ErrorMessage}");
        Assert.Equal("PieChart", readResult.Name);
        Assert.Equal("Sheet1", readResult.SheetName);
        Assert.Equal(ChartType.Pie, readResult.ChartType);
        Assert.False(readResult.IsPivotChart);
        Assert.True(readResult.Series.Count > 0);
    }

    [Fact]
    public void Read_NonExistentChart_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(Read_NonExistentChart_ReturnsError),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Read(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Delete_ExistingChart_RemovesChart()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(Delete_ExistingChart_RemovesChart),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B3"].Value2 = new object[,] { { "X", "Y" }, { 1, 10 }, { 2, 20 } };
            return 0;
        });

        _commands.CreateFromRange(batch, "Sheet1", "A1:B3", ChartType.Line, 50, 50);

        var listBefore = _commands.List(batch);
        Assert.Single(listBefore.Charts);
        string chartName = listBefore.Charts[0].Name;

        // Act
        var deleteResult = _commands.Delete(batch, chartName);

        // Assert
        Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");

        // Verify chart removed
        var listAfter = _commands.List(batch);
        Assert.True(listAfter.Success);
        Assert.Empty(listAfter.Charts);
    }

    [Fact]
    public void Move_ExistingChart_UpdatesPosition()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(Move_ExistingChart_UpdatesPosition),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B3"].Value2 = new object[,] { { "X", "Y" }, { 1, 10 }, { 2, 20 } };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B3", ChartType.ColumnClustered, 100, 100, 300, 200);
        Assert.True(createResult.Success);

        // Act
        var moveResult = _commands.Move(batch, createResult.ChartName, left: 200, top: 150, width: 400, height: 250);

        // Assert
        Assert.True(moveResult.Success, $"Move failed: {moveResult.ErrorMessage}");

        // Verify position updated
        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.True(readResult.Success);
        Assert.Equal(200, readResult.Left);
        Assert.Equal(150, readResult.Top);
        Assert.Equal(400, readResult.Width);
        Assert.Equal(250, readResult.Height);
    }

    [Fact]
    public void List_MultipleCharts_ReturnsAll()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(List_MultipleCharts_ReturnsAll),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "Cat", "Val" },
                { "A", 10 },
                { "B", 20 },
                { "C", 30 }
            };
            return 0;
        });

        // Create multiple charts
        _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50, 300, 200, "Chart1");
        _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Pie, 400, 50, 300, 200, "Chart2");
        _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Line, 50, 300, 300, 200, "Chart3");

        // Act
        var listResult = _commands.List(batch);

        // Assert
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        Assert.Equal(3, listResult.Charts.Count);
        Assert.Contains(listResult.Charts, c => c.Name == "Chart1" && c.ChartType == ChartType.ColumnClustered);
        Assert.Contains(listResult.Charts, c => c.Name == "Chart2" && c.ChartType == ChartType.Pie);
        Assert.Contains(listResult.Charts, c => c.Name == "Chart3" && c.ChartType == ChartType.Line);
    }

    [Fact]
    public void CreateFromRange_DifferentChartTypes_CreatesCorrectly()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(CreateFromRange_DifferentChartTypes_CreatesCorrectly),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:C5"].Value2 = new object[,] {
                { "Month", "Series1", "Series2" },
                { "Jan", 10, 20 },
                { "Feb", 15, 25 },
                { "Mar", 20, 30 },
                { "Apr", 25, 35 }
            };
            return 0;
        });

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
            Assert.True(result.Success, $"Failed to create {chartType}: {result.ErrorMessage}");
            Assert.Equal(chartType, result.ChartType);
            x += 300;
        }

        // Verify all created
        var listResult = _commands.List(batch);
        Assert.Equal(chartTypes.Length, listResult.Charts.Count);
    }
}
