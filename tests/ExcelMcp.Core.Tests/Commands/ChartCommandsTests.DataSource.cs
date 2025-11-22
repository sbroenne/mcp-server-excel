using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Chart data source operations (SetSourceRange, AddSeries, RemoveSeries).
/// </summary>
public partial class ChartCommandsTests
{
    [Fact]
    public void SetSourceRange_RegularChart_UpdatesDataSource()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(SetSourceRange_RegularChart_UpdatesDataSource),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create initial data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "Original", "Val" },
                { "A", 10 },
                { "B", 20 },
                { "C", 30 }
            };
            sheet.Range["D1:E5"].Value2 = new object[,] {
                { "New", "Data" },
                { "X", 100 },
                { "Y", 200 },
                { "Z", 300 },
                { "W", 400 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50);
        Assert.True(createResult.Success);

        // Act
        var setRangeResult = _commands.SetSourceRange(batch, createResult.ChartName, "D1:E5");

        // Assert
        Assert.True(setRangeResult.Success, $"SetSourceRange failed: {setRangeResult.ErrorMessage}");

        // Verify source range changed
        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.True(readResult.Success);
        Assert.Contains("D1:E5", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddSeries_RegularChart_AddsNewSeries()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(AddSeries_RegularChart_AddsNewSeries),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:C4"].Value2 = new object[,] {
                { "Cat", "Series1", "Series2" },
                { "Q1", 10, 15 },
                { "Q2", 20, 25 },
                { "Q3", 30, 35 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Line, 50, 50);
        Assert.True(createResult.Success);

        var readBefore = _commands.Read(batch, createResult.ChartName);
        int initialSeriesCount = readBefore.Series.Count;

        // Act
        var addSeriesResult = _commands.AddSeries(
            batch,
            createResult.ChartName,
            "NewSeries",
            "Sheet1!C2:C4",
            "Sheet1!A2:A4");

        // Assert
        Assert.True(addSeriesResult.Success, $"AddSeries failed: {addSeriesResult.ErrorMessage}");
        Assert.Equal("NewSeries", addSeriesResult.SeriesName);
        Assert.True(addSeriesResult.SeriesIndex > 0);

        // Verify series added
        var readAfter = _commands.Read(batch, createResult.ChartName);
        Assert.True(readAfter.Success);
        Assert.Equal(initialSeriesCount + 1, readAfter.Series.Count);
        Assert.Contains(readAfter.Series, s => s.Name == "NewSeries");
    }

    [Fact]
    public void RemoveSeries_RegularChart_RemovesSeries()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(RemoveSeries_RegularChart_RemovesSeries),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart with multiple series
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:C4"].Value2 = new object[,] {
                { "Cat", "Series1", "Series2" },
                { "A", 10, 20 },
                { "B", 15, 25 },
                { "C", 20, 30 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:C4", ChartType.ColumnClustered, 50, 50);
        Assert.True(createResult.Success);

        var readBefore = _commands.Read(batch, createResult.ChartName);
        int initialSeriesCount = readBefore.Series.Count;
        Assert.True(initialSeriesCount >= 2, "Need at least 2 series for test");

        // Act - Remove first series (index 1)
        var removeResult = _commands.RemoveSeries(batch, createResult.ChartName, 1);

        // Assert
        Assert.True(removeResult.Success, $"RemoveSeries failed: {removeResult.ErrorMessage}");

        // Verify series removed
        var readAfter = _commands.Read(batch, createResult.ChartName);
        Assert.True(readAfter.Success);
        Assert.Equal(initialSeriesCount - 1, readAfter.Series.Count);
    }

    [Fact]
    public void AddSeries_WithoutCategoryRange_CreatesSeriesSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(AddSeries_WithoutCategoryRange_CreatesSeriesSuccessfully),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 20 },
                { 3, 30 }
            };
            sheet.Range["C2:C4"].Value2 = new object[,] { { 15 }, { 25 }, { 35 } };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.XYScatter, 50, 50);
        Assert.True(createResult.Success);

        // Act - Add series without category range
        var addResult = _commands.AddSeries(batch, createResult.ChartName, "Series3", "Sheet1!C2:C4", null);

        // Assert
        Assert.True(addResult.Success, $"AddSeries failed: {addResult.ErrorMessage}");
        Assert.Equal("Series3", addResult.SeriesName);
    }

    [Fact]
    public void SetSourceRange_NonExistentChart_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(SetSourceRange_NonExistentChart_ReturnsError),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.SetSourceRange(batch, "NonExistent", "A1:B10");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddSeries_InvalidSeriesIndex_HandlesGracefully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(AddSeries_InvalidSeriesIndex_HandlesGracefully),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 20 },
                { 3, 30 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Line, 50, 50);
        Assert.True(createResult.Success);

        // Act - Try to remove non-existent series index
        var removeResult = _commands.RemoveSeries(batch, createResult.ChartName, 999);

        // Assert - Should fail gracefully
        Assert.False(removeResult.Success);
    }

    [Fact]
    public void SetSourceRange_ExpandedRange_UpdatesChartData()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ChartCommandsTests),
            nameof(SetSourceRange_ExpandedRange_UpdatesChartData),
            _tempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create initial small dataset
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B3"].Value2 = new object[,] {
                { "Cat", "Val" },
                { "A", 10 },
                { "B", 20 }
            };
            // Add more data rows
            sheet.Range["A4"].Value2 = "C";
            sheet.Range["B4"].Value2 = 30;
            sheet.Range["A5"].Value2 = "D";
            sheet.Range["B5"].Value2 = 40;
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B3", ChartType.BarClustered, 50, 50);
        Assert.True(createResult.Success);

        // Act - Expand to include more rows
        var expandResult = _commands.SetSourceRange(batch, createResult.ChartName, "A1:B5");

        // Assert
        Assert.True(expandResult.Success, $"SetSourceRange failed: {expandResult.ErrorMessage}");

        // Verify expanded range
        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.True(readResult.Success);
        Assert.Contains("A1:B5", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
    }
}
