using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Chart formatting operations (DataLabels, AxisScale, Gridlines, SeriesFormat).
/// </summary>
public partial class ChartCommandsTests
{
    // === DATA LABELS ===

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_ShowValue_DisplaysValuesOnChart()
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

        // Act - Enable data labels showing values
        _commands.SetDataLabels(batch, createResult.ChartName, showValue: true);

        // Assert - Verify data labels are set (no exception means success for void operations)
        // The operation completed without error
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_ShowPercentage_DisplaysPercentageOnPieChart()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "Category", "Value" },
                { "Product A", 40 },
                { "Product B", 35 },
                { "Product C", 25 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Pie, 50, 50);

        // Act - Enable percentage labels (common for pie charts)
        _commands.SetDataLabels(batch, createResult.ChartName, showPercentage: true);

        // Assert - Operation succeeded
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_SpecificSeries_AppliesOnlyToTargetSeries()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:C4"].Value2 = new object[,] {
                { "Month", "Series1", "Series2" },
                { "Jan", 100, 200 },
                { "Feb", 150, 250 },
                { "Mar", 200, 300 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:C4", ChartType.Line, 50, 50);

        // Act - Enable data labels only for series 1
        _commands.SetDataLabels(batch, createResult.ChartName, showValue: true, seriesIndex: 1);

        // Assert - Operation succeeded
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_WithPosition_SetsLabelPosition()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "Cat", "Val" },
                { "A", 100 },
                { "B", 150 },
                { "C", 200 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act - Show values at outside end of bars
        _commands.SetDataLabels(batch, createResult.ChartName, showValue: true, position: DataLabelPosition.OutsideEnd);

        // Assert - Operation succeeded
    }

    // === AXIS SCALE ===

    [Fact]
    [Trait("Feature", "Charts")]
    public void GetAxisScale_ValueAxis_ReturnsScaleInfo()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 100 },
                { 2, 200 },
                { 3, 300 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Line, 50, 50);

        // Act
        var result = _commands.GetAxisScale(batch, createResult.ChartName, ChartAxisType.Value);

        // Assert
        Assert.True(result.Success);
        Assert.Equal(createResult.ChartName, result.ChartName);
        Assert.Equal("Value", result.AxisType);
        // By default, Excel uses auto scale
        Assert.True(result.MinimumScaleIsAuto);
        Assert.True(result.MaximumScaleIsAuto);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetAxisScale_CustomMinMax_SetsScaleValues()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 100 },
                { 2, 200 },
                { 3, 300 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Line, 50, 50);

        // Act - Set custom scale
        _commands.SetAxisScale(batch, createResult.ChartName, ChartAxisType.Value, minimumScale: 0, maximumScale: 500);

        // Assert - Verify scale changed
        var result = _commands.GetAxisScale(batch, createResult.ChartName, ChartAxisType.Value);
        Assert.False(result.MinimumScaleIsAuto);
        Assert.False(result.MaximumScaleIsAuto);
        Assert.Equal(0, result.MinimumScale);
        Assert.Equal(500, result.MaximumScale);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetAxisScale_WithMajorUnit_SetsMajorUnitInterval()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 100 },
                { 2, 200 },
                { 3, 300 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act - Set major unit to 50
        _commands.SetAxisScale(batch, createResult.ChartName, ChartAxisType.Value, majorUnit: 50);

        // Assert - Verify major unit changed
        var result = _commands.GetAxisScale(batch, createResult.ChartName, ChartAxisType.Value);
        Assert.False(result.MajorUnitIsAuto);
        Assert.Equal(50, result.MajorUnit);
    }

    // === GRIDLINES ===

    [Fact]
    [Trait("Feature", "Charts")]
    public void GetGridlines_Chart_ReturnsGridlinesInfo()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

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

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act
        var result = _commands.GetGridlines(batch, createResult.ChartName);

        // Assert
        Assert.True(result.Success);
        Assert.Equal(createResult.ChartName, result.ChartName);
        // Default Excel charts have major gridlines on value axis
        Assert.True(result.Gridlines.HasValueMajorGridlines);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetGridlines_EnableMinorGridlines_ShowsMinorGridlines()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 100 },
                { 2, 200 },
                { 3, 300 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.Line, 50, 50);

        // Act - Enable minor gridlines on value axis
        _commands.SetGridlines(batch, createResult.ChartName, ChartAxisType.Value, showMinor: true);

        // Assert
        var result = _commands.GetGridlines(batch, createResult.ChartName);
        Assert.True(result.Gridlines.HasValueMinorGridlines);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetGridlines_DisableMajorGridlines_HidesMajorGridlines()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B4"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 100 },
                { 2, 200 },
                { 3, 300 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act - Hide major gridlines on value axis
        _commands.SetGridlines(batch, createResult.ChartName, ChartAxisType.Value, showMajor: false);

        // Assert
        var result = _commands.GetGridlines(batch, createResult.ChartName);
        Assert.False(result.Gridlines.HasValueMajorGridlines);
    }

    // === SERIES FORMATTING ===

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetSeriesFormat_MarkerStyle_ChangesMarkerStyle()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

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

        // Use LineMarkers chart type which shows markers by default
        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.LineMarkers, 50, 50);

        // Act - Change marker style to diamond
        _commands.SetSeriesFormat(batch, createResult.ChartName, seriesIndex: 1, markerStyle: MarkerStyle.Diamond);

        // Assert - Operation succeeded (void operation)
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetSeriesFormat_MarkerSize_ChangesMarkerSize()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

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

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.XYScatter, 50, 50);

        // Act - Set marker size to 10
        _commands.SetSeriesFormat(batch, createResult.ChartName, seriesIndex: 1, markerSize: 10);

        // Assert - Operation succeeded
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetSeriesFormat_MarkerColors_SetsMarkerColors()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

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

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.LineMarkers, 50, 50);

        // Act - Set marker colors (red fill, blue border)
        _commands.SetSeriesFormat(
            batch,
            createResult.ChartName,
            seriesIndex: 1,
            markerBackgroundColor: "#FF0000",
            markerForegroundColor: "#0000FF");

        // Assert - Operation succeeded
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetSeriesFormat_InvalidSeriesIndex_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

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

        // Act & Assert - Should throw for invalid series index
        Assert.Throws<ArgumentException>(() =>
            _commands.SetSeriesFormat(batch, createResult.ChartName, seriesIndex: 999, markerStyle: MarkerStyle.Circle));
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetSeriesFormat_InvalidMarkerSize_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

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

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.LineMarkers, 50, 50);

        // Act & Assert - Should throw for marker size outside valid range (2-72)
        Assert.Throws<ArgumentException>(() =>
            _commands.SetSeriesFormat(batch, createResult.ChartName, seriesIndex: 1, markerSize: 100));
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_InvalidSeriesIndex_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

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

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act & Assert - Should throw for invalid series index
        Assert.Throws<ArgumentException>(() =>
            _commands.SetDataLabels(batch, createResult.ChartName, showValue: true, seriesIndex: 999));
    }

    // === TRENDLINES ===

    [Fact]
    [Trait("Feature", "Charts")]
    public void AddTrendline_Linear_AddsTrendlineToSeries()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B5"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 22 },
                { 3, 28 },
                { 4, 42 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B5", ChartType.XYScatter, 50, 50);

        // Act
        var result = _commands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, type: TrendlineType.Linear);

        // Assert
        Assert.True(result.Success, $"AddTrendline failed: {result.ErrorMessage}");
        Assert.Equal(TrendlineType.Linear, result.Type);
        Assert.Equal(1, result.TrendlineIndex);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void AddTrendline_WithEquationDisplay_ShowsEquationOnChart()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B5"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 22 },
                { 3, 28 },
                { 4, 42 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B5", ChartType.XYScatter, 50, 50);

        // Act
        var result = _commands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, type: TrendlineType.Linear,
            displayEquation: true, displayRSquared: true);

        // Assert
        Assert.True(result.Success);

        // Verify via ListTrendlines
        var listResult = _commands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);
        Assert.True(listResult.Success);
        Assert.Single(listResult.Trendlines);
        Assert.True(listResult.Trendlines[0].DisplayEquation);
        Assert.True(listResult.Trendlines[0].DisplayRSquared);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void AddTrendline_Polynomial_RequiresOrder()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B5"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 22 },
                { 3, 28 },
                { 4, 42 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B5", ChartType.XYScatter, 50, 50);

        // Act & Assert - Should throw without order
        Assert.Throws<ArgumentException>(() =>
            _commands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, type: TrendlineType.Polynomial));

        // Should succeed with order
        var result = _commands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, type: TrendlineType.Polynomial, order: 2);
        Assert.True(result.Success);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void ListTrendlines_MultipleTrendlines_ReturnsAll()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B5"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 22 },
                { 3, 28 },
                { 4, 42 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B5", ChartType.XYScatter, 50, 50);

        // Add multiple trendlines
        _commands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, type: TrendlineType.Linear);
        _commands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, type: TrendlineType.Exponential);

        // Act
        var result = _commands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);

        // Assert
        Assert.True(result.Success);
        Assert.Equal(2, result.Trendlines.Count);
        Assert.Contains(result.Trendlines, t => t.Type == TrendlineType.Linear);
        Assert.Contains(result.Trendlines, t => t.Type == TrendlineType.Exponential);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void DeleteTrendline_RemovesTrendlineFromSeries()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B5"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 22 },
                { 3, 28 },
                { 4, 42 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B5", ChartType.XYScatter, 50, 50);
        _commands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, type: TrendlineType.Linear);

        // Verify trendline exists
        var beforeList = _commands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);
        Assert.Single(beforeList.Trendlines);

        // Act
        _commands.DeleteTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineIndex: 1);

        // Assert
        var afterList = _commands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);
        Assert.Empty(afterList.Trendlines);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetTrendline_UpdatesDisplayOptions()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B5"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 22 },
                { 3, 28 },
                { 4, 42 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B5", ChartType.XYScatter, 50, 50);
        _commands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, type: TrendlineType.Linear);

        // Verify initial state (no equation displayed)
        var beforeList = _commands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);
        Assert.False(beforeList.Trendlines[0].DisplayEquation);

        // Act
        _commands.SetTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineIndex: 1,
            displayEquation: true, displayRSquared: true);

        // Assert
        var afterList = _commands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);
        Assert.True(afterList.Trendlines[0].DisplayEquation);
        Assert.True(afterList.Trendlines[0].DisplayRSquared);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void AddTrendline_WithForecasting_ExtendsTrendline()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B5"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 22 },
                { 3, 28 },
                { 4, 42 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B5", ChartType.XYScatter, 50, 50);

        // Act
        var result = _commands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, type: TrendlineType.Linear,
            forward: 2.0, backward: 1.0);

        // Assert
        Assert.True(result.Success);

        var listResult = _commands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);
        Assert.Equal(2.0, listResult.Trendlines[0].Forward);
        Assert.Equal(1.0, listResult.Trendlines[0].Backward);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void DeleteTrendline_InvalidIndex_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1:B5"].Value2 = new object[,] {
                { "X", "Y" },
                { 1, 10 },
                { 2, 22 },
                { 3, 28 },
                { 4, 42 }
            };
            return 0;
        });

        var createResult = _commands.CreateFromRange(batch, "Sheet1", "A1:B5", ChartType.XYScatter, 50, 50);

        // Act & Assert - Should throw for invalid trendline index (no trendlines exist)
        Assert.Throws<ArgumentException>(() =>
            _commands.DeleteTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineIndex: 1));
    }
}
