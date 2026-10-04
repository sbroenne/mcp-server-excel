using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for Chart appearance operations (SetChartType, SetTitle, SetAxisTitle, ShowLegend, SetStyle, Get/SetAxisNumberFormat).
/// </summary>
public sealed partial class PersistentServiceChartFormattingTests
{
    [Theory]
    [InlineData("[Red]0.00")]
    [InlineData("[>=1.5]0.00;[Red]0.00")]
    [InlineData("General")]
    [InlineData("$#,##0.00")]
    [InlineData("mmm-yy")]
    [InlineData("yyyy-mm-dd hh:mm:ss")]
    [InlineData("0.00,,\"M\"")]
    [InlineData("[h]:mm:ss")]
    public void AxisNumberFormat_NativeInvariantColorControl_PreservesCodeAndDisplay(string format)
    {
        var created = RequireSuccess(_chartCommands.CreateFromRange(
            _fixture.BatchToken, _sheetName, "A1:B4", ChartType.ColumnClustered));
        var translated = _fixture.ExecuteRawVerification((context, _) =>
            context.FormatTranslator.TranslateForChart(format));
        var local = InspectChart(created.ChartName, chart =>
        {
            Microsoft.Office.Interop.Excel.Axis? axis = null;
            Microsoft.Office.Interop.Excel.TickLabels? labels = null;
            try
            {
                axis = (Microsoft.Office.Interop.Excel.Axis)chart.Axes(
                    Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
                labels = axis.TickLabels;
                labels.NumberFormat = translated;
                Assert.Equal(translated, labels.NumberFormat);
                return labels.NumberFormat;
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref labels);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref axis);
            }
        });
        if (!format.Contains("[Red]", StringComparison.Ordinal))
        {
            return;
        }
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Microsoft.Office.Interop.Excel.Sheets? sheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
            Microsoft.Office.Interop.Excel.Range? cell = null;
            Microsoft.Office.Interop.Excel.DisplayFormat? display = null;
            Microsoft.Office.Interop.Excel.Font? font = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[_sheetName];
                cell = sheet.Range["Z1"];
                cell.Value2 = 1.25;
                cell.NumberFormat = context.FormatTranslator.TranslateFromLocale(local);
                Assert.Equal($"1{context.FormatTranslator.DecimalSeparator}25", cell.Text);
                display = cell.DisplayFormat;
                font = display.Font;
                Assert.Equal(255, Convert.ToInt32(font.Color, System.Globalization.CultureInfo.InvariantCulture));
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref font);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref display);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref cell);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref sheet);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref sheets);
            }
        });
    }

    [Theory]
    [InlineData("[Red]$0.00")]
    [InlineData("[Red]\"$\"0.00")]
    public void AxisNumberFormat_NamedColorWithDollar_PreservesCurrency(string format)
    {
        var created = RequireSuccess(_chartCommands.CreateFromRange(
            _fixture.BatchToken, _sheetName, "A1:B4", ChartType.ColumnClustered));
        RequireSuccess(_chartCommands.SetAxisNumberFormat(
            _fixture.BatchToken, created.ChartName, ChartAxisType.Value, format));
        Assert.Equal("[Red]\\$0.00", _chartCommands.GetAxisNumberFormat(
            _fixture.BatchToken, created.ChartName, ChartAxisType.Value));
        var actualCode = InspectChart(created.ChartName, chart =>
        {
            Microsoft.Office.Interop.Excel.Axis? axis = null;
            Microsoft.Office.Interop.Excel.TickLabels? labels = null;
            try
            {
                axis = (Microsoft.Office.Interop.Excel.Axis)chart.Axes(
                    Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
                labels = axis.TickLabels;
                return labels.NumberFormat;
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref labels);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref axis);
            }
        });
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Microsoft.Office.Interop.Excel.Sheets? sheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
            Microsoft.Office.Interop.Excel.Range? cell = null;
            try
            {
                Assert.Equal($"[Red]\\$0{context.FormatTranslator.DecimalSeparator}00", actualCode);
                sheets = context.Book.Worksheets;
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[_sheetName];
                cell = sheet.Range["Z1"];
                cell.Value2 = 1.25;
                cell.NumberFormat = "[Red]\\$0.00";
                Assert.Equal($"$1{context.FormatTranslator.DecimalSeparator}25", cell.Text);
                Assert.Equal(1.25, Assert.IsType<double>(cell.Value2));
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref cell);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref sheet);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref sheets);
            }
        });
    }

    [Theory]
    [InlineData("$#,##0.00", 1234.56, "$1{group}234{decimal}56")]
    [InlineData("0.00,,\"M\"", 1250000, "1{decimal}25M")]
    [InlineData("[>=1.5]\"high\";\"low\"", 1.25, "low")]
    [InlineData("[>=1.5]\"high\";\"low\"", 1.75, "high")]
    [InlineData("yyyy-mm-dd hh:mm:ss", 45000.75, "2023-03-15 18:00:00")]
    [InlineData("[h]:mm:ss", 1.75, "42:00:00")]
    [InlineData("0.00 \"a,b.c\"\\m", 12.5, "12{decimal}50 a,b.cm")]
    public void AxisNumberFormat_LocalProperty_PreservesDisplayMeaning(
        string format, double value, string expectedText)
    {
        var batch = _fixture.BatchToken;
        var created = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered);
        RequireSuccess(created);
        var written = _chartCommands.SetAxisNumberFormat(batch, created.ChartName, ChartAxisType.Value, format);
        RequireSuccess(written);
        Assert.Equal(format, _chartCommands.GetAxisNumberFormat(batch, created.ChartName, ChartAxisType.Value));
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Microsoft.Office.Interop.Excel.Sheets? sheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
            Microsoft.Office.Interop.Excel.ChartObjects? objects = null;
            Microsoft.Office.Interop.Excel.ChartObject? chartObject = null;
            Microsoft.Office.Interop.Excel.Chart? chart = null;
            Microsoft.Office.Interop.Excel.Axis? axis = null;
            Microsoft.Office.Interop.Excel.TickLabels? labels = null;
            Microsoft.Office.Interop.Excel.Range? cell = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[_sheetName];
                objects = (Microsoft.Office.Interop.Excel.ChartObjects)sheet.ChartObjects();
                chartObject = objects.Item(created.ChartName);
                chart = chartObject.Chart;
                axis = (Microsoft.Office.Interop.Excel.Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
                labels = axis.TickLabels;
                Assert.Equal(format, ctx.FormatTranslator.TranslateFromLocale(labels.NumberFormatLocal));
                // Excel's own cell renderer checks the meaning of the actual local axis code.
                cell = sheet.Range["Z1"];
                cell.ColumnWidth = 40;
                cell.Value2 = value;
                cell.NumberFormatLocal = labels.NumberFormatLocal;
                Assert.Equal(expectedText
                    .Replace("{decimal}", ctx.FormatTranslator.DecimalSeparator, StringComparison.Ordinal)
                    .Replace("{group}", ctx.FormatTranslator.ThousandsSeparator, StringComparison.Ordinal),
                    cell.Text);
                Assert.Equal(value, Assert.IsType<double>(cell.Value2));
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref cell);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref labels);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref axis);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref chart);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref chartObject);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref objects);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref sheet);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref sheets);
            }
        });
    }

    [Fact]
    public void SetChartType_ExistingChart_ChangesType()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));
        Assert.Equal(ChartType.ColumnClustered, createResult.ChartType);

        // Act - Change to Line chart
        RequireSuccess(_chartCommands.SetChartType(batch, createResult.ChartName, ChartType.Line));

        // Assert - Verify type changed
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Equal(ChartType.Line, readResult.ChartType);
    }

    [Fact]
    public void SetTitle_ValidTitle_SetsChartTitle()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.Pie, 50, 50));

        // Act
        RequireSuccess(_chartCommands.SetTitle(batch, createResult.ChartName, "Sales by Quarter"));

        // Assert - Verify title set
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
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
        RequireSuccess(createResult);
        RequireSuccess(_chartCommands.SetTitle(
            batch,
            createResult.ChartName,
            "Initial Title"));
        Assert.Equal("Initial Title", RequireSuccess(_chartCommands.Read(batch, createResult.ChartName)).Title);

        RequireSuccess(_chartCommands.SetTitle(batch, createResult.ChartName, ""));

        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
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
        RequireSuccess(createResult);
        RequireSuccess(_chartCommands.SetPlacement(batch, createResult.ChartName, 2,
            printObject: false, locked: false, roundedCorners: true));
        var before = ReadChartObjectProperties(createResult.ChartName);

        var exception = Assert.Throws<ArgumentException>(() =>
            _chartCommands.SetPlacement(batch, createResult.ChartName, 5));

        Assert.Contains(
            "placement",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, ReadChartObjectProperties(createResult.ChartName));
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
        RequireSuccess(createResult);

        var setResult = _chartCommands.SetPlotOptions(
            batch,
            createResult.ChartName,
            plotBy: ChartPlotBy.Rows,
            displayBlanksAs: ChartDisplayBlanksAs.Zero,
            plotVisibleOnly: false);

        RequireSuccess(setResult);
        var getResult = _chartCommands.GetPlotOptions(
            batch,
            createResult.ChartName);
        RequireSuccess(getResult);
        Assert.Equal(ChartPlotBy.Rows, getResult.PlotBy);
        Assert.Equal(ChartDisplayBlanksAs.Zero, getResult.DisplayBlanksAs);
        Assert.False(getResult.PlotVisibleOnly);
    }

    [Fact]
    public void SetAxisTitle_CategoryAxis_SetsTitleSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        RequireSuccess(_chartCommands.SetAxisTitle(batch, createResult.ChartName, ChartAxisType.Category, "Months"));
        Assert.Equal("Months", ReadAxisTitle(createResult.ChartName, ChartAxisType.Category));
        Assert.Equal("", ReadAxisTitle(createResult.ChartName, ChartAxisType.Value));
    }

    [Fact]
    public void SetAxisTitle_ValueAxis_SetsTitleSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.BarClustered, 50, 50));

        RequireSuccess(_chartCommands.SetAxisTitle(batch, createResult.ChartName, ChartAxisType.Value, "Revenue ($)"));
        Assert.Equal("Revenue ($)", ReadAxisTitle(createResult.ChartName, ChartAxisType.Value));
        Assert.Equal("", ReadAxisTitle(createResult.ChartName, ChartAxisType.Category));
    }

    [Fact]
    public void ShowLegend_WithPosition_DisplaysLegendAtPosition()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create test data and chart
        RequireSuccess(_commands.SetValues(
            batch,
            _sheetName,
            "A1:C4",
            [["X", "Series1", "Series2"],
             ["A", 10, 20],
             ["B", 15, 25],
             ["C", 20, 30]], overwritePolicy: OverwritePolicy.Allow));

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:C4", ChartType.Line, 50, 50));

        // Act - Show legend at bottom
        RequireSuccess(_chartCommands.ShowLegend(batch, createResult.ChartName, true, LegendPosition.Bottom));

        // Assert - Verify legend visible
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.True(readResult.HasLegend);
        Assert.Equal((int)LegendPosition.Bottom, ReadLegendPosition(createResult.ChartName));
    }

    [Fact]
    public void ShowLegend_HideLegend_RemovesLegend()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.Area, 50, 50));
        RequireSuccess(_chartCommands.ShowLegend(batch, createResult.ChartName, true, LegendPosition.Right));
        Assert.True(RequireSuccess(_chartCommands.Read(batch, createResult.ChartName)).HasLegend);

        // Act - Hide legend
        RequireSuccess(_chartCommands.ShowLegend(batch, createResult.ChartName, false));

        // Assert - Verify legend hidden
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.False(readResult.HasLegend);
    }

    [Fact]
    public void SetStyle_ValidStyleId_AppliesStyle()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        RequireSuccess(_chartCommands.SetStyle(batch, createResult.ChartName, 10));
        Assert.Equal(10, InspectChart(createResult.ChartName,
            chart => Convert.ToInt32(chart.ChartStyle, System.Globalization.CultureInfo.InvariantCulture)));
    }

    [Fact]
    public void SetStyle_InvalidStyleId_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.Pie, 50, 50));
        var before = InspectChart(createResult.ChartName,
            chart => Convert.ToInt32(chart.ChartStyle, System.Globalization.CultureInfo.InvariantCulture));

        // Act & Assert - Invalid style ID should throw exception
        var exception = Assert.Throws<ArgumentException>(() =>
            _chartCommands.SetStyle(batch, createResult.ChartName, 999));
        Assert.Contains("between 1 and 48", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, InspectChart(createResult.ChartName,
            chart => Convert.ToInt32(chart.ChartStyle, System.Globalization.CultureInfo.InvariantCulture)));
    }

    [Fact]
    public void SetChartType_MultipleTypes_AllWorkCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B5", ChartType.ColumnClustered, 50, 50));

        // Act & Assert - Test multiple chart type changes
        var chartTypes = new[] { ChartType.Line, ChartType.Area, ChartType.BarClustered, ChartType.XYScatter, ChartType.Pie };

        foreach (var chartType in chartTypes)
        {
            RequireSuccess(_chartCommands.SetChartType(batch, createResult.ChartName, chartType));
            var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
            Assert.Equal(chartType, readResult.ChartType);
        }
    }

    [Fact]
    public void ShowLegend_DifferentPositions_AllWorkCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:C4", ChartType.ColumnClustered, 50, 50));

        // Act & Assert - Test all legend positions
        var positions = new[] {
            LegendPosition.Bottom,
            LegendPosition.Top,
            LegendPosition.Left,
            LegendPosition.Right,
            LegendPosition.Corner
        };

        foreach (var position in positions)
        {
            RequireSuccess(_chartCommands.ShowLegend(batch, createResult.ChartName, true, position));
            Assert.Equal((int)position, ReadLegendPosition(createResult.ChartName));
        }
    }

    [Fact]
    public void GetAxisNumberFormat_ValueAxis_ReturnsCurrentFormat()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        // Act
        RequireSuccess(createResult);
        RequireSuccess(_chartCommands.SetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value, "0.000"));
        var format = _chartCommands.GetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value);

        Assert.Equal("0.000", format);
    }

    [Fact]
    public void SetAxisNumberFormat_ValueAxis_SetsFormatSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        // Act - Set millions format
        RequireSuccess(_chartCommands.SetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value, "$#,##0,,\"M\""));

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
        RequireSuccess(_commands.SetValues(
            batch,
            _sheetName,
            "A1:B4",
            [["Date", "Sales"],
             [45658, 100],
             [45689, 150],
             [45717, 200]], overwritePolicy: OverwritePolicy.Allow));

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));

        // Act - Set date format on category axis
        RequireSuccess(_chartCommands.SetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Category, "mmm-yy"));

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
        RequireSuccess(_commands.SetValues(
            batch,
            _sheetName,
            "A1:B4",
            [["Item", "Rate"],
             ["A", 0.25],
             ["B", 0.50],
             ["C", 0.75]], overwritePolicy: OverwritePolicy.Allow));

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.BarClustered, 50, 50));

        // Act - Set percentage format
        RequireSuccess(_chartCommands.SetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value, "0%"));

        // Assert - Verify format was set
        var format = _chartCommands.GetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value);
        Assert.Equal("0%", format);
    }

    [Fact]
    public void SetAxisNumberFormat_NonExistentChart_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var before = CreateFormattingGuard();

        // Act & Assert - Non-existent chart should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.SetAxisNumberFormat(batch, "NonExistentChart", ChartAxisType.Value, "#,##0"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertFormattingGuard(before);
    }

    [Fact]
    public void GetAxisNumberFormat_NonExistentChart_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var before = CreateFormattingGuard();

        // Act & Assert - Non-existent chart should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.GetAxisNumberFormat(batch, "NonExistentChart", ChartAxisType.Value));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertFormattingGuard(before);
    }

    [Fact]
    public void SetAxisNumberFormat_MultipleFormats_AllWorkCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        // Act & Assert - Test multiple format changes
        var formats = new[] { "#,##0", "$#,##0", "#,##0.00", "$#,##0,,\"M\"", "0.0E+0" };

        foreach (var fmt in formats)
        {
            RequireSuccess(_chartCommands.SetAxisNumberFormat(batch, createResult.ChartName, ChartAxisType.Value, fmt));
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
        var chart = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        RequireSuccess(_chartCommands.SetAxisNumberFormat(batch, chart.ChartName, ChartAxisType.Value, format));

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

        RequireSuccess(chart);
        RequireSuccess(_chartCommands.SetAxisNumberFormat(
            batch,
            chart.ChartName,
            ChartAxisType.Value,
            "0.00 \"Total\""));

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

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        // Act - Set placement to MoveAndSize (1 = xlMoveAndSizeWithCells)
        RequireSuccess(_chartCommands.SetPlacement(batch, createResult.ChartName, 3));
        Assert.Equal(3, ReadChartObjectProperties(createResult.ChartName).Placement);
        RequireSuccess(_chartCommands.SetPlacement(batch, createResult.ChartName, 1));

        // Assert - Verify placement changed
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Equal(1, readResult.Placement);
    }

    [Fact]
    public void SetPlacement_MoveOnly_SetsPlacementMode()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));

        // Act - Set placement to Move (2 = xlMove)
        RequireSuccess(_chartCommands.SetPlacement(batch, createResult.ChartName, 3));
        Assert.Equal(3, ReadChartObjectProperties(createResult.ChartName).Placement);
        RequireSuccess(_chartCommands.SetPlacement(batch, createResult.ChartName, 2));

        // Assert - Verify placement changed
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Equal(2, readResult.Placement);
    }

    [Fact]
    public void SetPlacement_FreeFloating_SetsPlacementMode()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Pie, 50, 50));

        // Act - Set placement to FreeFloating (3 = xlFreeFloating)
        RequireSuccess(_chartCommands.SetPlacement(batch, createResult.ChartName, 1));
        Assert.Equal(1, ReadChartObjectProperties(createResult.ChartName).Placement);
        RequireSuccess(_chartCommands.SetPlacement(batch, createResult.ChartName, 3));

        // Assert - Verify placement changed
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Equal(3, readResult.Placement);
    }

    [Fact]
    public void SetPlacement_NonExistentChart_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var before = CreateFormattingGuard();

        // Act & Assert - Non-existent chart should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.SetPlacement(batch, "NonExistentChart", 1));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertFormattingGuard(before);
    }

    // === FIT TO RANGE TESTS ===

    [Fact]
    public void FitToRange_ValidRange_ResizesChartToMatchRange()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50, 400, 300));

        // Act - Fit chart to a specific range
        RequireSuccess(_chartCommands.FitToRange(batch, createResult.ChartName, _sheetName, "E5:J15"));

        AssertFitsRange(createResult.ChartName, _sheetName, "E5:J15");
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Equal("$E$5", readResult.TopLeftCell);
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    [Fact]
    public void FitToRange_DifferentSheet_UsesTargetGeometryWithoutMovingChartSheet()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create test data, chart, and a second sheet
        var targetSheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetColumnWidth(batch, targetSheet, "E:H", 20));
        RequireSuccess(_commands.SetRowHeight(batch, targetSheet, "1:10", 25));

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        RequireSuccess(_chartCommands.FitToRange(batch, createResult.ChartName, targetSheet, "E5:H10"));

        AssertFitsRange(createResult.ChartName, targetSheet, "E5:H10");
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    [Fact]
    public void FitToRange_NonExistentChart_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var before = CreateFormattingGuard();

        // Act & Assert - Non-existent chart should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.FitToRange(batch, "NonExistentChart", _sheetName, "A1:D10"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertFormattingGuard(before);
    }

    [Fact]
    public void FitToRange_InvalidRangeAddress_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));
        var before = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));

        // Act & Assert - Invalid range should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.FitToRange(batch, createResult.ChartName, _sheetName, "InvalidRange!!!"));
        Assert.Contains("chart.fit-to-range failed [ComInterop/COMException]", exception.Message);
        AssertChartUnchanged(before);
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    // === ANCHOR CELLS TESTS ===

    [Fact]
    public void Read_ChartCreatedAtPosition_ReturnsAnchorCells()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create chart at position left=50, top=50
        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50, 400, 300));

        // Act
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));

        Assert.Equal(50, readResult.Left);
        Assert.Equal(50, readResult.Top);
        Assert.Equal(400, readResult.Width);
        Assert.Equal(300, readResult.Height);
        AssertAnchorCells(readResult);
    }

    [Fact]
    public void Read_ChartAfterFitToRange_ReturnsUpdatedAnchorCells()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));

        // Get initial anchor cells
        var initialRead = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        AssertAnchorCells(initialRead);
        var initialTopLeft = initialRead.TopLeftCell;

        // Act - Fit chart to a different range
        RequireSuccess(_chartCommands.FitToRange(batch, createResult.ChartName, _sheetName, "F10:K20"));

        // Assert - Anchor cells should have changed
        var afterRead = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.NotEqual(initialTopLeft, afterRead.TopLeftCell);

        Assert.Equal("$F$10", afterRead.TopLeftCell);
        AssertAnchorCells(afterRead);
        AssertFitsRange(createResult.ChartName, _sheetName, "F10:K20");
    }

    [Fact]
    public void Read_ChartWithPlacement_ReturnsPlacementValue()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));
        RequireSuccess(_chartCommands.SetPlacement(batch, createResult.ChartName, 3));

        // Act
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));

        Assert.Equal(3, readResult.Placement);
        Assert.Equal(3, ReadChartObjectProperties(createResult.ChartName).Placement);
    }
}
