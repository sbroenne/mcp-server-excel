using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceChartFormattingTests
{
    [Theory]
    [InlineData("Left", 1)]
    [InlineData("Center", 2)]
    [InlineData("Right", 3)]
    public void SetTitle_AppliesPositionAndParagraphAlignment(string alignment, int expectedAlignment)
    {
        var chartName = CreateFormattingChart();
        _fixture.Send("chartconfig.set-title", new { chartName, title = "Before" });
        var before = ReadTitleLayout(chartName);

        var response = _fixture.Send("chartconfig.set-title", new
        {
            chartName,
            title = "Revenue",
            left = 16.0,
            top = 12.0,
            alignment
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var actual = ReadTitleLayout(chartName);
        Assert.Equal("Revenue", actual.Text);
        Assert.Equal(16, actual.Left, precision: 1);
        Assert.Equal(12, actual.Top, precision: 1);
        Assert.Equal((Office.MsoParagraphAlignment)expectedAlignment, actual.Alignment);
        Assert.NotEqual(before.Left, actual.Left);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("top")]
    [InlineData("alignment")]
    public void SetTitle_PreservesOmittedLayoutSettings(string setting)
    {
        var chartName = CreateFormattingChart();
        _fixture.Send("chartconfig.set-title", new
        {
            chartName,
            title = "Before",
            left = 16.0,
            top = 12.0,
            alignment = "Right"
        });
        var before = ReadTitleLayout(chartName);
        var args = new Dictionary<string, object?> { ["chartName"] = chartName, ["title"] = "Revenue" };
        args[setting] = setting == "alignment" ? "Left" : 24.0;
        _fixture.Send("chartconfig.set-title", args);
        var actual = ReadTitleLayout(chartName);
        Assert.Equal("Revenue", actual.Text);
        Assert.Equal(setting == "left" ? 24 : before.Left, actual.Left, precision: 1);
        Assert.Equal(setting == "top" ? 24 : before.Top, actual.Top, precision: 1);
        Assert.Equal(setting == "alignment" ? Office.MsoParagraphAlignment.msoAlignLeft : before.Alignment, actual.Alignment);
        _fixture.Send("chartconfig.set-title", new { chartName, title = "Another" });
        var textOnly = ReadTitleLayout(chartName);
        Assert.Equal("Another", textOnly.Text);
        Assert.Equal(actual.Left, textOnly.Left);
        Assert.Equal(actual.Top, textOnly.Top);
        Assert.Equal(actual.Alignment, textOnly.Alignment);
        _fixture.Send("chartconfig.set-title", new { chartName, title = "" });
        Assert.False(InspectChart(chartName, chart => chart.HasTitle));
    }

    [Theory]
    [InlineData("Category", 1, 1)]
    [InlineData("Value", 2, 1)]
    [InlineData("CategorySecondary", 1, 2)]
    [InlineData("ValueSecondary", 2, 2)]
    [InlineData("Primary", 1, 1)]
    [InlineData("Secondary", 2, 1)]
    public void SetAxisFormat_AppliesLineSettings(string axis, int type, int group)
    {
        var chartName = CreateFormattingChart();
        _fixture.Send("chartconfig.set-series-axis-group", new { chartName, seriesIndex = 2, axisGroup = "Secondary" });
        InspectChart(chartName, chart =>
        {
            chart.HasAxis[Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlSecondary] = true;
            return true;
        });
        var otherType = type == 1 ? Excel.XlAxisType.xlValue : Excel.XlAxisType.xlCategory;
        var untouched = ReadAxisLine(chartName, otherType);
        var response = _fixture.Send("chartconfig.set-axis-format", new
        {
            chartName,
            axis,
            lineColor = "#D0D0D0",
            lineWeight = 2.5,
            lineVisible = true
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var actual = ReadAxisLine(chartName, (Excel.XlAxisType)type, (Excel.XlAxisGroup)group);
        Assert.Equal(0xD0D0D0, actual.Color);
        Assert.Equal(2.5f, actual.Weight);
        Assert.Equal(Office.MsoTriState.msoTrue, actual.Visible);
        Assert.Equal(untouched, ReadAxisLine(chartName, otherType));
        _fixture.Send("chartconfig.set-axis-format", new { chartName, axis, lineVisible = false });
        var hidden = ReadAxisLine(chartName, (Excel.XlAxisType)type, (Excel.XlAxisGroup)group);
        Assert.Equal(Office.MsoTriState.msoFalse, hidden.Visible);
        Assert.Equal(actual.Color, hidden.Color);
        Assert.Equal(actual.Weight, hidden.Weight);
        _fixture.Send("chartconfig.set-axis-format", new { chartName, axis, lineColor = "#2468AC" });
        var recolored = ReadAxisLine(chartName, (Excel.XlAxisType)type, (Excel.XlAxisGroup)group);
        Assert.Equal(0xAC6824, recolored.Color);
        Assert.Equal(hidden.Weight, recolored.Weight);
        Assert.Equal(hidden.Visible, recolored.Visible);
        _fixture.Send("chartconfig.set-axis-format", new { chartName, axis, lineWeight = 1.5 });
        var resized = ReadAxisLine(chartName, (Excel.XlAxisType)type, (Excel.XlAxisGroup)group);
        Assert.Equal(recolored.Color, resized.Color);
        Assert.Equal(1.5f, resized.Weight);
        Assert.Equal(hidden.Visible, resized.Visible);
        _fixture.Send("chartconfig.set-axis-format", new { chartName, axis, lineVisible = true });
        Assert.Equal((resized.Color, resized.Weight, Office.MsoTriState.msoTrue),
            ReadAxisLine(chartName, (Excel.XlAxisType)type, (Excel.XlAxisGroup)group));
    }

    [Theory]
    [InlineData("Chart")]
    [InlineData("Plot")]
    public void SetAreaFormat_HidesBorderWithoutChangingItsOtherSettings(string area)
    {
        var chartName = CreateFormattingChart();
        _fixture.Send("chartconfig.set-area-format", new
        {
            chartName,
            area,
            lineColor = "#2468AC",
            lineWeight = 2.5
        });
        var before = ReadAreaLine(chartName, area);
        Assert.Equal(Office.MsoTriState.msoTrue, before.Visible);
        var response = _fixture.Send("chartconfig.set-area-format", new { chartName, area, lineVisible = false });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var actual = ReadAreaLine(chartName, area);
        Assert.Equal(Office.MsoTriState.msoFalse, actual.Visible);
        Assert.Equal(before.Color, actual.Color);
        Assert.Equal(before.Weight, actual.Weight);
        _fixture.Send("chartconfig.set-area-format", new { chartName, area, fillColor = "#FFFFFF" });
        Assert.Equal(actual, ReadAreaLine(chartName, area));
        _fixture.Send("chartconfig.set-area-format", new
        {
            chartName,
            area,
            lineColor = "#D0D0D0",
            lineWeight = 1.5,
            lineVisible = false
        });
        Assert.Equal((0xD0D0D0, 1.5f, Office.MsoTriState.msoFalse), ReadAreaLine(chartName, area));
        _fixture.Send("chartconfig.set-area-format", new { chartName, area, lineVisible = true });
        Assert.Equal((0xD0D0D0, 1.5f, Office.MsoTriState.msoTrue), ReadAreaLine(chartName, area));
        _fixture.Send("chartconfig.set-area-format", new { chartName, area, lineVisible = false });
        _fixture.Send("chartconfig.set-area-format", new { chartName, area, lineWeight = 2.0 });
        Assert.Equal((0xD0D0D0, 2.0f, Office.MsoTriState.msoTrue), ReadAreaLine(chartName, area));
    }

    [Theory]
    [InlineData("set-title", """{"title":"","left":10}""", "hiding")]
    [InlineData("set-title", """{"title":"Changed","alignment":"Unknown"}""", "alignment")]
    [InlineData("set-axis-format", """{"axis":"Unknown","lineVisible":false}""", "axis")]
    [InlineData("set-axis-format", """{"axis":"Value"}""", "at least one")]
    [InlineData("set-axis-format", """{"axis":"Value","lineWeight":0}""", "greater than zero")]
    [InlineData("set-axis-format", """{"axis":"Value","lineWeight":-1,"lineColor":"#FF0000"}""", "greater than zero")]
    [InlineData("set-axis-format", """{"axis":"Value","lineColor":"red","lineVisible":false}""", "color")]
    [InlineData("set-area-format", """{"area":"Unknown","lineVisible":false}""", "area")]
    [InlineData("set-area-format", """{"area":"Chart","fillColor":"#FF0000","lineWeight":0,"lineVisible":false}""", "greater than zero")]
    public async Task TitleAndLineSettings_InvalidInputPreservesChart(string action, string json, string message)
    {
        var chartName = CreateFormattingChart();
        _fixture.Send("chartconfig.set-title", new { chartName, title = "Before", left = 16.0, top = 12.0, alignment = "Right" });
        var title = ReadTitleLayout(chartName);
        var axis = ReadAxisLine(chartName, Excel.XlAxisType.xlValue);
        var area = ReadAreaLine(chartName, "Chart");
        var args = System.Text.Json.Nodes.JsonNode.Parse(json)!.AsObject();
        args["chartName"] = chartName;
        var failure = await _fixture.SendForFailureAsync($"chartconfig.{action}", args);
        Assert.False(failure.Success);
        Assert.Contains(message, failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(title, ReadTitleLayout(chartName));
        Assert.Equal(axis, ReadAxisLine(chartName, Excel.XlAxisType.xlValue));
        Assert.Equal(area, ReadAreaLine(chartName, "Chart"));
    }

    [Theory]
    [InlineData("Value")]
    [InlineData("Category")]
    [InlineData("ValueSecondary")]
    public async Task SetAxisFormat_MissingAxisFailsWithoutChangingChart(string axis)
    {
        var chartName = CreateFormattingChart();
        _fixture.Send("chartconfig.set-chart-type", new { chartName, chartType = "Pie" });
        var before = _fixture.Send("chart.read", new { chartName }).Result;
        var failure = await _fixture.SendForFailureAsync("chartconfig.set-axis-format", new { chartName, axis, lineVisible = false });
        Assert.False(failure.Success);
        Assert.Contains($"no {axis} axis", failure.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(before, _fixture.Send("chart.read", new { chartName }).Result);
    }

    [Fact]
    public async Task TitleAndLineSettings_PersistAfterSaveAndReopen()
    {
        var chartName = CreateFormattingChart();
        _fixture.Send("chartconfig.set-title", new { chartName, title = "Revenue", left = 16.0, top = 12.0, alignment = "Left" });
        _fixture.Send("chartconfig.set-axis-format", new { chartName, axis = "Value", lineColor = "#D0D0D0", lineWeight = 2.5, lineVisible = false });
        _fixture.Send("chartconfig.set-area-format", new { chartName, area = "Chart", lineColor = "#2468AC", lineWeight = 1.5, lineVisible = false });
        var title = ReadTitleLayout(chartName);
        var axis = ReadAxisLine(chartName, Excel.XlAxisType.xlValue);
        var area = ReadAreaLine(chartName, "Chart");
        await _fixture.SaveAndReopenAsync();
        Assert.Equal(title, ReadTitleLayout(chartName));
        Assert.Equal(axis, ReadAxisLine(chartName, Excel.XlAxisType.xlValue));
        Assert.Equal(area, ReadAreaLine(chartName, "Chart"));
    }

    [Fact]
    public async Task SetTitle_ProtectedChartReportsFailure()
    {
        var chartName = CreateFormattingChart();
        _fixture.Send("chartconfig.set-title", new { chartName, title = "Before", left = 16.0, top = 12.0, alignment = "Right" });
        var title = ReadTitleLayout(chartName);
        var axis = ReadAxisLine(chartName, Excel.XlAxisType.xlValue);
        var area = ReadAreaLine(chartName, "Chart");
        _fixture.Send("worksheetstyle.set-protection", new { sheetName = _sheetName, isProtected = true });
        var failure = await Record.ExceptionAsync(async () =>
        {
            var response = await _fixture.SendForFailureAsync("chartconfig.set-title", new { chartName, title = "Changed", left = 24.0 });
            Assert.False(response.Success);
            Assert.False(string.IsNullOrWhiteSpace(response.HResult));
            Assert.Contains(chartName, response.ErrorMessage, StringComparison.Ordinal);
            Assert.Contains("no rollback", response.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal(title, ReadTitleLayout(chartName));
            Assert.Equal(axis, ReadAxisLine(chartName, Excel.XlAxisType.xlValue));
            Assert.Equal(area, ReadAreaLine(chartName, "Chart"));
        });
        var cleanup = Record.Exception(() => _fixture.Send("worksheetstyle.set-protection", new { sheetName = _sheetName, isProtected = false }));
        if (cleanup is not null) failure = PersistentServiceCleanupFailures.Combine(failure, cleanup);
        PersistentServiceCleanupFailures.ThrowIfFailed(failure);
    }

    private string CreateFormattingChart() =>
        RequireSuccess(_chartCommands.CreateFromRange(_fixture.BatchToken, _sheetName, "A1:C6",
            ChartType.ColumnClustered, chartName: $"Layout_{Guid.NewGuid():N}")).ChartName;

    private (string Text, double Left, double Top, Office.MsoParagraphAlignment Alignment)
        ReadTitleLayout(string chartName) =>
        InspectChart(chartName, chart =>
        {
            Excel.ChartTitle? title = null;
            Excel.ChartFormat? format = null;
            Excel.TextFrame2? frame = null;
            Office.TextRange2? text = null;
            Office.ParagraphFormat2? paragraph = null;
            try
            {
                Assert.True(chart.HasTitle);
                title = chart.ChartTitle;
                format = title.Format;
                frame = format.TextFrame2;
                text = frame.TextRange;
                paragraph = text.ParagraphFormat;
                return (title.Text, title.Left, title.Top, paragraph.Alignment);
            }
            finally
            {
                ComUtilities.Release(ref paragraph);
                ComUtilities.Release(ref text);
                ComUtilities.Release(ref frame);
                ComUtilities.Release(ref format);
                ComUtilities.Release(ref title);
            }
        });

    private (int Color, float Weight, Office.MsoTriState Visible) ReadAxisLine(
        string chartName, Excel.XlAxisType type, Excel.XlAxisGroup group = Excel.XlAxisGroup.xlPrimary) =>
        InspectChart(chartName, chart =>
        {
            Excel.Axis? axis = null;
            Excel.ChartFormat? format = null;
            try
            {
                axis = (Excel.Axis)chart.Axes(type, group);
                format = axis.Format;
                return ReadLine(format);
            }
            finally
            {
                ComUtilities.Release(ref format);
                ComUtilities.Release(ref axis);
            }
        });

    private (int Color, float Weight, Office.MsoTriState Visible) ReadAreaLine(string chartName, string area) =>
        InspectChart(chartName, chart =>
        {
            Excel.ChartArea? chartArea = null;
            Excel.PlotArea? plotArea = null;
            Excel.ChartFormat? format = null;
            try
            {
                if (area == "Chart") { chartArea = chart.ChartArea; format = chartArea.Format; }
                else { plotArea = chart.PlotArea; format = plotArea.Format; }
                return ReadLine(format);
            }
            finally
            {
                ComUtilities.Release(ref format);
                ComUtilities.Release(ref plotArea);
                ComUtilities.Release(ref chartArea);
            }
        });

    private static (int Color, float Weight, Office.MsoTriState Visible) ReadLine(Excel.ChartFormat format)
    {
        Excel.LineFormat? line = null;
        Excel.ColorFormat? color = null;
        try
        {
            line = format.Line;
            color = line.ForeColor;
            return (color.RGB, line.Weight, line.Visible);
        }
        finally
        {
            ComUtilities.Release(ref color);
            ComUtilities.Release(ref line);
        }
    }
}
