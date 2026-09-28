using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for Chart appearance operations (SetChartType, SetTitle, SetAxisTitle, ShowLegend, SetStyle, Get/SetAxisNumberFormat).
/// </summary>
public sealed partial class PersistentServiceChartFormattingTests
{
    [Fact]
    public void SetChartType_ExistingChart_ChangesType()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);
        Assert.Equal(ChartType.ColumnClustered, createResult.ChartType);

        // Act - Change to Line chart
        _chartCommands.SetChartType(batch, createResult.ChartName, ChartType.Line);

        // Assert - Verify type changed
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.Equal(ChartType.Line, readResult.ChartType);
    }

    [Fact]
    public void SetTitle_ValidTitle_SetsChartTitle()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.Pie, 50, 50);

        // Act
        _chartCommands.SetTitle(batch, createResult.ChartName, "Sales by Quarter");

        // Assert - Verify title set
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.Equal("Sales by Quarter", readResult.Title);
    }

    [Fact]
    public void SetTitle_EmptyString_HidesTitle()
    {
        var batch = _fixture.BatchToken;
        var createResult = _chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A1:B3",
            ChartType.BarClustered,
            50,
            50);
        _chartCommands.SetTitle(
            batch,
            createResult.ChartName,
            "Initial Title");

        _chartCommands.SetTitle(batch, createResult.ChartName, "");

        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.Null(readResult.Title);
    }

    [Fact]
    public void SetPlacement_InvalidValue_ThrowsHelpfulPublicError()
    {
        var batch = _fixture.BatchToken;
        var createResult = _chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A1:B3",
            ChartType.ColumnClustered,
            50,
            50);

        var exception = Assert.Throws<ArgumentException>(() =>
            _chartCommands.SetPlacement(batch, createResult.ChartName, 5));

        Assert.Contains(
            "placement",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PlotOptions_SetAndGet_RoundTripsChartBehavior()
    {
        var batch = _fixture.BatchToken;
        var createResult = _chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A1:C6",
            ChartType.Line,
            chartName: $"Plot_{Guid.NewGuid():N}"[..20]);

        var setResult = _chartCommands.SetPlotOptions(
            batch,
            createResult.ChartName,
            plotBy: ChartPlotBy.Rows,
            displayBlanksAs: ChartDisplayBlanksAs.Zero,
            plotVisibleOnly: false);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        var getResult = _chartCommands.GetPlotOptions(
            batch,
            createResult.ChartName);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(ChartPlotBy.Rows, getResult.PlotBy);
        Assert.Equal(ChartDisplayBlanksAs.Zero, getResult.DisplayBlanksAs);
        Assert.False(getResult.PlotVisibleOnly);
    }

    [Fact]
    public void SetAxisTitle_CategoryAxis_SetsTitleSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act & Assert - void operation, no exception means success
        _chartCommands.SetAxisTitle(batch, createResult.ChartName, ChartAxisType.Category, "Months");
    }

    [Fact]
    public void SetAxisTitle_ValueAxis_SetsTitleSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.BarClustered, 50, 50);

        // Act & Assert - void operation, no exception means success
        _chartCommands.SetAxisTitle(batch, createResult.ChartName, ChartAxisType.Value, "Revenue ($)");
    }

    [Fact]
    public void ShowLegend_WithPosition_DisplaysLegendAtPosition()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create test data and chart
        _commands.SetValues(
            batch,
            _sheetName,
            "A1:C4",
            [["X", "Series1", "Series2"],
             ["A", 10, 20],
             ["B", 15, 25],
             ["C", 20, 30]]);

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:C4", ChartType.Line, 50, 50);

        // Act - Show legend at bottom
        _chartCommands.ShowLegend(batch, createResult.ChartName, true, LegendPosition.Bottom);

        // Assert - Verify legend visible
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.True(readResult.HasLegend);
    }

    [Fact]
    public void ShowLegend_HideLegend_RemovesLegend()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.Area, 50, 50);
        _chartCommands.ShowLegend(batch, createResult.ChartName, true, LegendPosition.Right); // Show first

        // Act - Hide legend
        _chartCommands.ShowLegend(batch, createResult.ChartName, false);

        // Assert - Verify legend hidden
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.False(readResult.HasLegend);
    }

    [Fact]
    public void SetStyle_ValidStyleId_AppliesStyle()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act & Assert - void operation, no exception means success
        _chartCommands.SetStyle(batch, createResult.ChartName, 10);
    }

    [Fact]
    public void SetStyle_InvalidStyleId_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.Pie, 50, 50);

        // Act & Assert - Invalid style ID should throw exception
        var exception = Assert.Throws<ArgumentException>(() =>
            _chartCommands.SetStyle(batch, createResult.ChartName, 999));
        Assert.Contains("between 1 and 48", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetChartType_MultipleTypes_AllWorkCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B5", ChartType.ColumnClustered, 50, 50);

        // Act & Assert - Test multiple chart type changes
        var chartTypes = new[] { ChartType.Line, ChartType.Area, ChartType.BarClustered, ChartType.XYScatter, ChartType.Pie };

        foreach (var chartType in chartTypes)
        {
            _chartCommands.SetChartType(batch, createResult.ChartName, chartType);
            var readResult = _chartCommands.Read(batch, createResult.ChartName);
            Assert.Equal(chartType, readResult.ChartType);
        }
    }

    [Fact]
    public void ShowLegend_DifferentPositions_AllWorkCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:C4", ChartType.ColumnClustered, 50, 50);

        // Act & Assert - Test all legend positions
        var positions = new[] {
            LegendPosition.Bottom,
            LegendPosition.Top,
            LegendPosition.Left,
            LegendPosition.Right,
            LegendPosition.Corner
        };

        foreach (var position in positions)
            _chartCommands.ShowLegend(batch, createResult.ChartName, true, position);
    }

    [Fact]
    public void GetAxisNumberFormat_ValueAxis_ReturnsCurrentFormat()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act
        var format = _chartCommands.GetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value);

        // Assert - Default format is typically "General"
        Assert.NotNull(format);
    }

    [Fact]
    public void SetAxisNumberFormat_ValueAxis_SetsFormatSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act - Set millions format
        _chartCommands.SetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value, "$#,##0,,\"M\"");

        // Assert - Verify format was set
        var format = _chartCommands.GetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value);
        Assert.Equal("$#,##0,,\"M\"", format);
    }

    [Fact]
    public void SetAxisNumberFormat_CategoryAxis_SetsFormatSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create test data with dates
        _commands.SetValues(
            batch,
            _sheetName,
            "A1:B4",
            [["Date", "Sales"],
             [45658, 100],
             [45689, 150],
             [45717, 200]]);

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 50);

        // Act - Set date format on category axis
        _chartCommands.SetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Category, "mmm-yy");

        // Assert - Verify format was set
        var format = _chartCommands.GetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Category);
        Assert.Equal("mmm-yy", format);
    }

    [Fact]
    public void SetAxisNumberFormat_PercentageFormat_SetsFormatSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create test data
        _commands.SetValues(
            batch,
            _sheetName,
            "A1:B4",
            [["Item", "Rate"],
             ["A", 0.25],
             ["B", 0.50],
             ["C", 0.75]]);

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.BarClustered, 50, 50);

        // Act - Set percentage format
        _chartCommands.SetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value, "0%");

        // Assert - Verify format was set
        var format = _chartCommands.GetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value);
        Assert.Equal("0%", format);
    }

    [Fact]
    public void SetAxisNumberFormat_NonExistentChart_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act & Assert - Non-existent chart should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.SetAxisNumberFormat(batch, "NonExistentChart", ChartAxisType.Value, "#,##0"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetAxisNumberFormat_NonExistentChart_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act & Assert - Non-existent chart should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.GetAxisNumberFormat(batch, "NonExistentChart", ChartAxisType.Value));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetAxisNumberFormat_MultipleFormats_AllWorkCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act & Assert - Test multiple format changes
        var formats = new[] { "#,##0", "$#,##0", "#,##0.00", "$#,##0,,\"M\"", "0.0E+0" };

        foreach (var fmt in formats)
        {
            _chartCommands.SetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value, fmt);
            var result = _chartCommands.GetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value);
            Assert.Equal(fmt, result);
        }
    }

    [Theory]
    [InlineData("General")]
    [InlineData("[Red]0.00")]
    [InlineData("[>=1.5]0.00;[Red]0.00")]
    public void SetAxisNumberFormat_KeywordsAndConditions_RoundTripsInvariantCode(string format)
    {
        var batch = _fixture.BatchToken;
        var chart = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        _chartCommands.SetAxisNumberFormat(batch, chart.ChartName, ChartAxisType.Value, format);

        Assert.Equal(format, _chartCommands.GetAxisNumberFormat(batch, chart.ChartName, ChartAxisType.Value));
    }

    [Fact]
    public void SetAxisNumberFormat_LocalizedDateLettersInLiteral_RoundTripsInvariantCode()
    {
        var batch = _fixture.BatchToken;
        var chart = _chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A1:B4",
            ChartType.ColumnClustered,
            50,
            50);

        _chartCommands.SetAxisNumberFormat(
            batch,
            chart.ChartName,
            ChartAxisType.Value,
            "0.00 \"Total\"");

        Assert.Equal(
            "0.00 \"Total\"",
            _chartCommands.GetAxisNumberFormat(batch, chart.ChartName, ChartAxisType.Value));
    }

    // === PLACEMENT TESTS ===

    [Fact]
    public void SetPlacement_MoveAndSize_SetsPlacementMode()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act - Set placement to MoveAndSize (1 = xlMoveAndSizeWithCells)
        _chartCommands.SetPlacement(batch, createResult.ChartName, 1);

        // Assert - Verify placement changed
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.Equal(1, readResult.Placement);
    }

    [Fact]
    public void SetPlacement_MoveOnly_SetsPlacementMode()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 50);

        // Act - Set placement to Move (2 = xlMove)
        _chartCommands.SetPlacement(batch, createResult.ChartName, 2);

        // Assert - Verify placement changed
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.Equal(2, readResult.Placement);
    }

    [Fact]
    public void SetPlacement_FreeFloating_SetsPlacementMode()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Pie, 50, 50);

        // Act - Set placement to FreeFloating (3 = xlFreeFloating)
        _chartCommands.SetPlacement(batch, createResult.ChartName, 3);

        // Assert - Verify placement changed
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.Equal(3, readResult.Placement);
    }

    [Fact]
    public void SetPlacement_NonExistentChart_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act & Assert - Non-existent chart should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.SetPlacement(batch, "NonExistentChart", 1));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    // === FIT TO RANGE TESTS ===

    [Fact]
    public void FitToRange_ValidRange_ResizesChartToMatchRange()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50, 400, 300);

        // Act - Fit chart to a specific range
        _chartCommands.FitToRange(batch, createResult.ChartName, _sheetName, "E5:J15");

        // Assert - Verify chart position/size changed (can verify via Read that chart still exists)
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.NotNull(readResult);

        // TopLeftCell should now be E5 (or close to it, depending on exact positioning)
        Assert.NotNull(readResult.TopLeftCell);
    }

    [Fact]
    public void FitToRange_DifferentSheet_ThrowsOrMovesChart()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create test data, chart, and a second sheet
        _fixture.CreateNamedTestSheet(batch, "Sheet2");

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act - Try to fit chart to range on same sheet
        _chartCommands.FitToRange(batch, createResult.ChartName, _sheetName, "E5:H10");

        // Assert - Verify chart moved
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.NotNull(readResult);
    }

    [Fact]
    public void FitToRange_NonExistentChart_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act & Assert - Non-existent chart should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.FitToRange(batch, "NonExistentChart", _sheetName, "A1:D10"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FitToRange_InvalidRangeAddress_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act & Assert - Invalid range should throw
        Assert.ThrowsAny<Exception>(() =>
            _chartCommands.FitToRange(batch, createResult.ChartName, _sheetName, "InvalidRange!!!"));
    }

    // === ANCHOR CELLS TESTS ===

    [Fact]
    public void Read_ChartCreatedAtPosition_ReturnsAnchorCells()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create chart at position left=50, top=50
        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50, 400, 300);

        // Act
        var readResult = _chartCommands.Read(batch, createResult.ChartName);

        // Assert - Anchor cells should be populated
        Assert.NotNull(readResult.TopLeftCell);
        Assert.NotNull(readResult.BottomRightCell);

        // TopLeftCell should be a valid cell address (e.g., "$A$1", "$B$2", etc.)
        Assert.Matches(@"\$[A-Z]+\$\d+", readResult.TopLeftCell);
        Assert.Matches(@"\$[A-Z]+\$\d+", readResult.BottomRightCell);
    }

    [Fact]
    public void Read_ChartAfterFitToRange_ReturnsUpdatedAnchorCells()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Get initial anchor cells
        var initialRead = _chartCommands.Read(batch, createResult.ChartName);
        var initialTopLeft = initialRead.TopLeftCell;

        // Act - Fit chart to a different range
        _chartCommands.FitToRange(batch, createResult.ChartName, _sheetName, "F10:K20");

        // Assert - Anchor cells should have changed
        var afterRead = _chartCommands.Read(batch, createResult.ChartName);
        Assert.NotEqual(initialTopLeft, afterRead.TopLeftCell);

        // The new TopLeftCell should reflect the new position (around F10)
        Assert.NotNull(afterRead.TopLeftCell);
        Assert.NotNull(afterRead.BottomRightCell);
    }

    [Fact]
    public void Read_ChartWithPlacement_ReturnsPlacementValue()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act
        var readResult = _chartCommands.Read(batch, createResult.ChartName);

        // Assert - Placement should be populated with a valid value (1, 2, or 3)
        Assert.NotNull(readResult.Placement);
        Assert.InRange(readResult.Placement.Value, 1, 3);
    }
}
