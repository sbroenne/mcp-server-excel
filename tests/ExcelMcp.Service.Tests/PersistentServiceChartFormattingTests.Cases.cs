using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for Chart formatting operations (DataLabels, AxisScale, Gridlines, SeriesFormat).
/// </summary>
public sealed partial class PersistentServiceChartFormattingTests
{
    // === DATA LABELS ===

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_ShowValue_DisplaysValuesOnChart()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        // Act - Enable data labels showing values
        RequireSuccess(_chartCommands.SetDataLabels(batch, createResult.ChartName, showValue: true));

        var state = ReadSeriesState(createResult.ChartName);
        Assert.True(state.HasLabels);
        Assert.True(state.ShowValue);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_ShowPercentage_DisplaysPercentageOnPieChart()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Pie, 50, 50));

        // Act - Enable percentage labels (common for pie charts)
        RequireSuccess(_chartCommands.SetDataLabels(batch, createResult.ChartName, showPercentage: true));

        var state = ReadSeriesState(createResult.ChartName);
        Assert.True(state.HasLabels);
        Assert.True(state.ShowPercentage);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_SpecificSeries_AppliesOnlyToTargetSeries()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:C4", ChartType.Line, 50, 50));

        var untouched = ReadSeriesState(createResult.ChartName, 2);
        // Act - Enable data labels only for series 1
        RequireSuccess(_chartCommands.SetDataLabels(batch, createResult.ChartName, showValue: true, seriesIndex: 1));

        Assert.True(ReadSeriesState(createResult.ChartName).ShowValue);
        Assert.Equal(untouched, ReadSeriesState(createResult.ChartName, 2));
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_WithPosition_SetsLabelPosition()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        // Act - Show values at outside end of bars
        RequireSuccess(_chartCommands.SetDataLabels(batch, createResult.ChartName, showValue: true, labelPosition: DataLabelPosition.OutsideEnd));

        var state = ReadSeriesState(createResult.ChartName);
        Assert.True(state.ShowValue);
        Assert.Equal((int)DataLabelPosition.OutsideEnd, state.LabelPosition);
    }

    /// <summary>
    /// Regression test: SetDataLabels with InsideEnd/InsideBase/OutsideEnd on Line charts
    /// must throw a friendly InvalidOperationException, not a raw COMException.
    /// These positions are only valid for bar/column/area chart types.
    /// </summary>
    [Theory]
    [InlineData(ChartType.Line, DataLabelPosition.InsideEnd)]
    [InlineData(ChartType.Line, DataLabelPosition.InsideBase)]
    [InlineData(ChartType.Line, DataLabelPosition.OutsideEnd)]
    [InlineData(ChartType.LineMarkers, DataLabelPosition.InsideEnd)]
    [InlineData(ChartType.LineMarkers, DataLabelPosition.InsideBase)]
    [InlineData(ChartType.LineMarkers, DataLabelPosition.OutsideEnd)]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_InsideEndOnLineChart_ThrowsFriendlyException(
        ChartType chartType, DataLabelPosition position)
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", chartType, 50, 50));
        var before = ReadSeriesState(createResult.ChartName);

        // Act & Assert - must throw InvalidOperationException (not COMException)
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.SetDataLabels(batch, createResult.ChartName, showValue: true, labelPosition: position));

        Assert.Contains(position.ToString(), ex.Message);
        Assert.Contains("not supported", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, ReadSeriesState(createResult.ChartName));
    }

    [Fact]
    public void SetDataLabels_InvalidPositionOnLaterSeries_PreservesEverySeries()
    {
        var batch = _fixture.BatchToken;
        var chart = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:C4", ChartType.ColumnClustered));
        RequireSuccess(_chartCommands.SetSeriesChartType(batch, chart.ChartName, 2, ChartType.Line));
        var first = ReadSeriesState(chart.ChartName);
        var second = ReadSeriesState(chart.ChartName, 2);

        var error = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.SetDataLabels(batch, chart.ChartName,
                showValue: true, labelPosition: DataLabelPosition.OutsideEnd));

        Assert.Contains("OutsideEnd", error.Message);
        Assert.Equal(first, ReadSeriesState(chart.ChartName));
        Assert.Equal(second, ReadSeriesState(chart.ChartName, 2));
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_AboveOnLineChart_Succeeds()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));

        // Act - Above is valid for line charts
        RequireSuccess(_chartCommands.SetDataLabels(batch, createResult.ChartName, showValue: true, labelPosition: DataLabelPosition.Above));

        var state = ReadSeriesState(createResult.ChartName);
        Assert.True(state.ShowValue);
        Assert.Equal((int)DataLabelPosition.Above, state.LabelPosition);
    }

    // === AXIS SCALE ===

    [Fact]
    [Trait("Feature", "Charts")]
    public void GetAxisScale_ValueAxis_ReturnsScaleInfo()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));

        // Act
        var result = RequireSuccess(_chartCommands.GetAxisScale(batch, createResult.ChartName, ChartAxisType.Value));

        // Assert
        RequireSuccess(result);
        Assert.Equal(createResult.ChartName, result.ChartName);
        Assert.Equal("Value", result.AxisType);
        // By default, Excel uses auto scale
        Assert.True(result.MinimumScaleIsAuto);
        Assert.True(result.MaximumScaleIsAuto);
        AssertNativeAxisScale(result);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetAxisScale_CustomMinMax_SetsScaleValues()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));

        // Act - Set custom scale
        RequireSuccess(_chartCommands.SetAxisScale(batch, createResult.ChartName, ChartAxisType.Value, minimumScale: 0, maximumScale: 500));

        // Assert - Verify scale changed
        var result = RequireSuccess(_chartCommands.GetAxisScale(batch, createResult.ChartName, ChartAxisType.Value));
        Assert.False(result.MinimumScaleIsAuto);
        Assert.False(result.MaximumScaleIsAuto);
        Assert.Equal(0, result.MinimumScale);
        Assert.Equal(500, result.MaximumScale);
        AssertNativeAxisScale(result);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetAxisScale_WithMajorUnit_SetsMajorUnitInterval()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        // Act - Set major unit to 50
        RequireSuccess(_chartCommands.SetAxisScale(batch, createResult.ChartName, ChartAxisType.Value, majorUnit: 50));

        // Assert - Verify major unit changed
        var result = RequireSuccess(_chartCommands.GetAxisScale(batch, createResult.ChartName, ChartAxisType.Value));
        Assert.False(result.MajorUnitIsAuto);
        Assert.Equal(50, result.MajorUnit);
        AssertNativeAxisScale(result);
    }

    // === GRIDLINES ===

    [Fact]
    [Trait("Feature", "Charts")]
    public void GetGridlines_Chart_ReturnsGridlinesInfo()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        // Act
        var result = RequireSuccess(_chartCommands.GetGridlines(batch, createResult.ChartName));

        // Assert
        RequireSuccess(result);
        Assert.Equal(createResult.ChartName, result.ChartName);
        // Default Excel charts have major gridlines on value axis
        Assert.True(result.Gridlines.HasValueMajorGridlines);
        AssertNativeGridlines(result);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetGridlines_EnableMinorGridlines_ShowsMinorGridlines()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));

        // Act - Enable minor gridlines on value axis
        RequireSuccess(_chartCommands.SetGridlines(batch, createResult.ChartName, ChartAxisType.Value, showMinor: true));

        // Assert
        var result = RequireSuccess(_chartCommands.GetGridlines(batch, createResult.ChartName));
        Assert.True(result.Gridlines.HasValueMinorGridlines);
        AssertNativeGridlines(result);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetGridlines_DisableMajorGridlines_HidesMajorGridlines()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        // Act - Hide major gridlines on value axis
        RequireSuccess(_chartCommands.SetGridlines(batch, createResult.ChartName, ChartAxisType.Value, showMajor: false));

        // Assert
        var result = RequireSuccess(_chartCommands.GetGridlines(batch, createResult.ChartName));
        Assert.False(result.Gridlines.HasValueMajorGridlines);
        AssertNativeGridlines(result);
    }

    // === SERIES FORMATTING ===

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetSeriesFormat_MarkerStyle_ChangesMarkerStyle()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Use LineMarkers chart type which shows markers by default
        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.LineMarkers, 50, 50));

        // Act - Change marker style to diamond
        RequireSuccess(_chartCommands.SetSeriesFormat(batch, createResult.ChartName, seriesIndex: 1, markerStyle: MarkerStyle.Diamond));

        Assert.Equal((int)MarkerStyle.Diamond, ReadSeriesState(createResult.ChartName).MarkerStyle);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetSeriesFormat_MarkerSize_ChangesMarkerSize()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.XYScatter, 50, 50));

        // Act - Set marker size to 10
        RequireSuccess(_chartCommands.SetSeriesFormat(batch, createResult.ChartName, seriesIndex: 1, markerSize: 10));

        Assert.Equal(10, ReadSeriesState(createResult.ChartName).MarkerSize);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetSeriesFormat_MarkerColors_SetsMarkerColors()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.LineMarkers, 50, 50));

        // Act - Set marker colors (red fill, blue border)
        RequireSuccess(_chartCommands.SetSeriesFormat(
            batch,
            createResult.ChartName,
            seriesIndex: 1,
            markerBackgroundColor: "#FF0000",
            markerForegroundColor: "#0000FF"));

        var state = ReadSeriesState(createResult.ChartName);
        Assert.Equal(0x0000FF, state.MarkerFill);
        Assert.Equal(0xFF0000, state.MarkerBorder);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetSeriesFormat_InvalidSeriesIndex_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));
        var before = ReadSeriesState(createResult.ChartName);

        // Act & Assert - Should throw for invalid series index
        Assert.Throws<ArgumentException>(() =>
            _chartCommands.SetSeriesFormat(batch, createResult.ChartName, seriesIndex: 999, markerStyle: MarkerStyle.Circle));
        Assert.Equal(before, ReadSeriesState(createResult.ChartName));
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetSeriesFormat_InvalidMarkerSize_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.LineMarkers, 50, 50));
        var before = ReadSeriesState(createResult.ChartName);

        // Act & Assert - Should throw for marker size outside valid range (2-72)
        Assert.Throws<ArgumentException>(() =>
            _chartCommands.SetSeriesFormat(batch, createResult.ChartName, seriesIndex: 1, markerSize: 100));
        Assert.Equal(before, ReadSeriesState(createResult.ChartName));
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetDataLabels_InvalidSeriesIndex_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));
        var before = ReadSeriesState(createResult.ChartName);

        // Act & Assert - Should throw for invalid series index
        Assert.Throws<ArgumentException>(() =>
            _chartCommands.SetDataLabels(batch, createResult.ChartName, showValue: true, seriesIndex: 999));
        Assert.Equal(before, ReadSeriesState(createResult.ChartName));
    }

    // === TRENDLINES ===

    [Fact]
    [Trait("Feature", "Charts")]
    public void AddTrendline_Linear_AddsTrendlineToSeries()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B5", ChartType.XYScatter, 50, 50));

        // Act
        var result = _chartCommands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineType: TrendlineType.Linear);

        // Assert
        RequireSuccess(result);
        Assert.Equal(TrendlineType.Linear, result.Type);
        Assert.Equal(1, result.TrendlineIndex);
        var listed = _chartCommands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);
        RequireSuccess(listed);
        Assert.Equal(TrendlineType.Linear, Assert.Single(listed.Trendlines).Type);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void AddTrendline_WithEquationDisplay_ShowsEquationOnChart()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B5", ChartType.XYScatter, 50, 50));

        // Act
        var result = _chartCommands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineType: TrendlineType.Linear,
            displayEquation: true, displayRSquared: true);

        // Assert
        RequireSuccess(result);

        // Verify via ListTrendlines
        var listResult = _chartCommands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);
        RequireSuccess(listResult);
        Assert.Single(listResult.Trendlines);
        Assert.True(listResult.Trendlines[0].DisplayEquation);
        Assert.True(listResult.Trendlines[0].DisplayRSquared);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void AddTrendline_Polynomial_RequiresOrder()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B5", ChartType.XYScatter, 50, 50));

        // Act & Assert - Should throw without order
        Assert.Throws<ArgumentException>(() =>
            _chartCommands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineType: TrendlineType.Polynomial));
        var rejected = _chartCommands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);
        RequireSuccess(rejected);
        Assert.Empty(rejected.Trendlines);

        // Should succeed with order
        var result = _chartCommands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineType: TrendlineType.Polynomial, order: 2);
        RequireSuccess(result);
        var listed = _chartCommands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);
        RequireSuccess(listed);
        var trendline = Assert.Single(listed.Trendlines);
        Assert.Equal(TrendlineType.Polynomial, trendline.Type);
        Assert.Equal(2, trendline.Order);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void ListTrendlines_MultipleTrendlines_ReturnsAll()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B5", ChartType.XYScatter, 50, 50));

        // Add multiple trendlines
        RequireSuccess(_chartCommands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineType: TrendlineType.Linear));
        RequireSuccess(_chartCommands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineType: TrendlineType.Exponential));

        // Act
        var result = _chartCommands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);

        // Assert
        RequireSuccess(result);
        Assert.Equal(2, result.Trendlines.Count);
        Assert.Contains(result.Trendlines, t => t.Type == TrendlineType.Linear);
        Assert.Contains(result.Trendlines, t => t.Type == TrendlineType.Exponential);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void DeleteTrendline_RemovesTrendlineFromSeries()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B5", ChartType.XYScatter, 50, 50));
        RequireSuccess(_chartCommands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineType: TrendlineType.Linear));

        // Verify trendline exists
        var beforeList = RequireSuccess(_chartCommands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1));
        Assert.Single(beforeList.Trendlines);

        // Act
        RequireSuccess(_chartCommands.DeleteTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineIndex: 1));

        // Assert
        var afterList = RequireSuccess(_chartCommands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1));
        Assert.Empty(afterList.Trendlines);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void SetTrendline_UpdatesDisplayOptions()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B5", ChartType.XYScatter, 50, 50));
        RequireSuccess(_chartCommands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineType: TrendlineType.Linear));

        // Verify initial state (no equation displayed)
        var beforeList = RequireSuccess(_chartCommands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1));
        Assert.Single(beforeList.Trendlines);
        Assert.False(beforeList.Trendlines[0].DisplayEquation);

        // Act
        RequireSuccess(_chartCommands.SetTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineIndex: 1,
            displayEquation: true, displayRSquared: true));

        // Assert
        var afterList = RequireSuccess(_chartCommands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1));
        Assert.True(afterList.Trendlines[0].DisplayEquation);
        Assert.True(afterList.Trendlines[0].DisplayRSquared);
        Assert.Single(afterList.Trendlines);
        Assert.Equal(TrendlineType.Linear, afterList.Trendlines[0].Type);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void AddTrendline_WithForecasting_ExtendsTrendline()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B5", ChartType.XYScatter, 50, 50));

        // Act
        var result = _chartCommands.AddTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineType: TrendlineType.Linear,
            forward: 2.0, backward: 1.0);

        // Assert
        RequireSuccess(result);

        var listResult = RequireSuccess(_chartCommands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1));
        Assert.Single(listResult.Trendlines);
        Assert.Equal(2.0, listResult.Trendlines[0].Forward);
        Assert.Equal(1.0, listResult.Trendlines[0].Backward);
    }

    [Fact]
    [Trait("Feature", "Charts")]
    public void DeleteTrendline_InvalidIndex_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B5", ChartType.XYScatter, 50, 50));

        // Act & Assert - Should throw for invalid trendline index (no trendlines exist)
        Assert.Throws<ArgumentException>(() =>
            _chartCommands.DeleteTrendline(batch, createResult.ChartName, seriesIndex: 1, trendlineIndex: 1));
        var listed = _chartCommands.ListTrendlines(batch, createResult.ChartName, seriesIndex: 1);
        RequireSuccess(listed);
        Assert.Empty(listed.Trendlines);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Trendline_InvalidIndex_PreservesExistingTrendline(bool delete)
    {
        var batch = _fixture.BatchToken;
        var created = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B5", ChartType.XYScatter);
        RequireSuccess(created);
        var added = _chartCommands.AddTrendline(batch, created.ChartName, 1, TrendlineType.Linear,
            displayEquation: true, forward: 2);
        RequireSuccess(added);

        Assert.Throws<ArgumentException>(() =>
        {
            if (delete)
                _chartCommands.DeleteTrendline(batch, created.ChartName, 1, 2);
            else
                _chartCommands.SetTrendline(batch, created.ChartName, 1, 2, displayEquation: false, forward: 4);
        });

        var listed = _chartCommands.ListTrendlines(batch, created.ChartName, 1);
        RequireSuccess(listed);
        var retained = Assert.Single(listed.Trendlines);
        Assert.Equal(TrendlineType.Linear, retained.Type);
        Assert.True(retained.DisplayEquation);
        Assert.Equal(2, retained.Forward);
    }
}
