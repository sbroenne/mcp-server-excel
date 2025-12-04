using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Chart appearance operations (SetChartType, SetTitle, SetAxisTitle, ShowLegend, SetStyle).
/// </summary>
public partial class ChartCommandsTests
{
    [Fact]
    public void SetChartType_ExistingChart_ChangesType()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

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

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50);
        Assert.Equal(ChartType.ColumnClustered, createResult.ChartType);

        // Act - Change to Line chart
        _commands.SetChartType(batch, createResult.ChartName, ChartType.Line);

        // Assert - Verify type changed
        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.Equal(ChartType.Line, readResult.ChartType);
    }

    [Fact]
    public void SetTitle_ValidTitle_SetsChartTitle()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B3"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 20 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B3", ChartType.Pie, 50, 50);

        // Act
        _commands.SetTitle(batch, createResult.ChartName, "Sales by Quarter");

        // Assert - Verify title set
        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.Equal("Sales by Quarter", readResult.Title);
    }

    [Fact]
    public void SetTitle_EmptyString_HidesTitle()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart with title
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B3"].Value2 = new object[,] { { "X", "Y" }, { 1, 10 }, { 2, 20 } };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B3", ChartType.BarClustered, 50, 50);
        _commands.SetTitle(batch, createResult.ChartName, "Initial Title");

        // Act - Hide title with empty string
        _commands.SetTitle(batch, createResult.ChartName, "");

        // Assert - Verify title hidden
        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.Null(readResult.Title);
    }

    [Fact]
    public void SetAxisTitle_CategoryAxis_SetsTitleSuccessfully()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "Month", "Sales" },
                { "Jan", 100 },
                { "Feb", 150 },
                { "Mar", 200 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act & Assert - void operation, no exception means success
        _commands.SetAxisTitle(batch, createResult.ChartName, ChartAxisType.Category, "Months");
    }

    [Fact]
    public void SetAxisTitle_ValueAxis_SetsTitleSuccessfully()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "Product", "Revenue" },
                { "A", 1000 },
                { "B", 1500 },
                { "C", 2000 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.BarClustered, 50, 50);

        // Act & Assert - void operation, no exception means success
        _commands.SetAxisTitle(batch, createResult.ChartName, ChartAxisType.Value, "Revenue ($)");
    }

    [Fact]
    public void ShowLegend_WithPosition_DisplaysLegendAtPosition()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:C4"].Value2 = new object[,] {
                { "X", "Series1", "Series2" },
                { "A", 10, 20 },
                { "B", 15, 25 },
                { "C", 20, 30 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:C4", ChartType.Line, 50, 50);

        // Act - Show legend at bottom
        _commands.ShowLegend(batch, createResult.ChartName, true, LegendPosition.Bottom);

        // Assert - Verify legend visible
        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.True(readResult.HasLegend);
    }

    [Fact]
    public void ShowLegend_HideLegend_RemovesLegend()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B3"].Value2 = new object[,] { { "X", "Y" }, { 1, 10 }, { 2, 20 } };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B3", ChartType.Area, 50, 50);
        _commands.ShowLegend(batch, createResult.ChartName, true, LegendPosition.Right); // Show first

        // Act - Hide legend
        _commands.ShowLegend(batch, createResult.ChartName, false);

        // Assert - Verify legend hidden
        var readResult = _commands.Read(batch, createResult.ChartName);
        Assert.False(readResult.HasLegend);
    }

    [Fact]
    public void SetStyle_ValidStyleId_AppliesStyle()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
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

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act & Assert - void operation, no exception means success
        _commands.SetStyle(batch, createResult.ChartName, 10);
    }

    [Fact]
    public void SetStyle_InvalidStyleId_ReturnsError()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B3"].Value2 = new object[,] { { "X", "Y" }, { 1, 10 }, { 2, 20 } };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B3", ChartType.Pie, 50, 50);

        // Act & Assert - Invalid style ID should throw exception
        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.SetStyle(batch, createResult.ChartName, 999));
        Assert.Contains("between 1 and 48", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetChartType_MultipleTypes_AllWorkCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B5"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 20 },
                { 3, 30 },
                { 4, 40 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B5", ChartType.ColumnClustered, 50, 50);

        // Act & Assert - Test multiple chart type changes
        var chartTypes = new[] { ChartType.Line, ChartType.Area, ChartType.BarClustered, ChartType.XYScatter, ChartType.Pie };

        foreach (var chartType in chartTypes)
        {
            _commands.SetChartType(batch, createResult.ChartName, chartType);
            var readResult = _commands.Read(batch, createResult.ChartName);
            Assert.Equal(chartType, readResult.ChartType);
        }
    }

    [Fact]
    public void ShowLegend_DifferentPositions_AllWorkCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create test data and chart
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:C4"].Value2 = new object[,] {
                { "X", "S1", "S2" },
                { "A", 10, 20 },
                { "B", 15, 25 },
                { "C", 20, 30 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:C4", ChartType.ColumnClustered, 50, 50);

        // Act & Assert - Test all legend positions
        var positions = new[] {
            LegendPosition.Bottom,
            LegendPosition.Top,
            LegendPosition.Left,
            LegendPosition.Right,
            LegendPosition.Corner
        };

        foreach (var position in positions)
            _commands.ShowLegend(batch, createResult.ChartName, true, position);
    }
}
