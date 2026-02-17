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
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        // Write additional data range for source change test
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["E1:F5"].Value2 = new object[,] {
                { "New", "Data" },
                { "X", 100 },
                { "Y", 200 },
                { "Z", 300 },
                { "W", 400 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act
        _commands.SetSourceRange(batch, createResult.ChartName, "E1:F5");

        // Assert - Verify source range changed (Excel returns SERIES formula with Sheet1 reference)
        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.Contains("Sheet1", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("$E$", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("$F$", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddSeries_RegularChart_AddsNewSeries()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Line, 50, 50);

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
        Assert.Equal("NewSeries", addSeriesResult.Name);

        // Verify series added
        var readAfter = _commands.Read(batch, createResult.ChartName);
        Assert.Equal(initialSeriesCount + 1, readAfter.Series.Count);
        Assert.Contains(readAfter.Series, s => s.Name == "NewSeries");
    }

    [Fact]
    public void RemoveSeries_RegularChart_RemovesSeries()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:C4", ChartType.ColumnClustered, 50, 50);

        var readBefore = _commands.Read(batch, createResult.ChartName);
        int initialSeriesCount = readBefore.Series.Count;
        Assert.True(initialSeriesCount >= 2, "Need at least 2 series for test");

        // Act - Remove first series (index 1)
        _commands.RemoveSeries(batch, createResult.ChartName, 1);

        // Assert - Verify series removed
        var readAfter = _commands.Read(batch, createResult.ChartName);
        Assert.Equal(initialSeriesCount - 1, readAfter.Series.Count);
    }

    [Fact]
    public void AddSeries_WithoutCategoryRange_CreatesSeriesSuccessfully()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.XYScatter, 50, 50);

        // Act - Add series without category range
        var addResult = _commands.AddSeries(batch, createResult.ChartName, "Series3", "Sheet1!C2:C4", null);

        // Assert
        Assert.Equal("Series3", addResult.Name);
    }

    [Fact]
    public void SetSourceRange_NonExistentChart_ReturnsError()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        // Act & Assert
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _commands.SetSourceRange(batch, "NonExistent", "A1:B10"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddSeries_InvalidSeriesIndex_HandlesGracefully()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Line, 50, 50);

        // Act & Assert - Invalid series index should throw exception
        var exception = Assert.Throws<System.Runtime.InteropServices.COMException>(() =>
            _commands.RemoveSeries(batch, createResult.ChartName, 999));

        Assert.NotNull(exception);
    }

    [Fact]
    public void SetSourceRange_ExpandedRange_UpdatesChartData()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B3", ChartType.BarClustered, 50, 50);

        // Act - Expand to include more rows
        _commands.SetSourceRange(batch, createResult.ChartName, "A1:B5");

        // Assert - Verify expanded range (Excel returns SERIES formula with Sheet1 reference)
        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.Contains("Sheet1", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("$A$", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        // Verify it includes row 5 (expanded from 3 to 5)
        Assert.Contains("$5", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
    }
}




