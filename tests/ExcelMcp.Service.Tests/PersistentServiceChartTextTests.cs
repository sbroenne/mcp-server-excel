using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceChartTextTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Theory]
    [InlineData("ChartArea", null)]
    [InlineData("Title", null)]
    [InlineData("CategoryAxis", null)]
    [InlineData("ValueAxis", null)]
    [InlineData("DataLabels", null)]
    [InlineData("Legend", null)]
    [InlineData("AxisTitle", "Value")]
    public void TextFormat_SetsAndReadsExistingElements(string element, string? axis)
    {
        var (sheet, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-title", new { chartName, title = "Revenue" });
        _fixture.Send("chartconfig.set-axis-title", new { chartName, axis = "Value", title = "Amount" });
        _fixture.Send("chartconfig.show-legend", new { chartName, visible = true });
        _fixture.Send("chartconfig.set-data-labels", new { chartName, showValue = true });
        var response = _fixture.Send("chartconfig.set-text-format", new
        {
            chartName,
            element,
            axis,
            fontName = "Arial",
            fontSize = 14.5,
            bold = true,
            italic = true,
            color = "#2468AC"
        });
        using var written = JsonDocument.Parse(response.Result!);
        AssertFormats(written.RootElement, element == "DataLabels" ? 2 : 1);
        using var read = JsonDocument.Parse(_fixture.Send("chartconfig.get-text-format", new { chartName, element, axis }).Result!);
        AssertFormats(read.RootElement, element == "DataLabels" ? 2 : 1);
        Assert.Equal(chartName, read.RootElement.GetProperty("chartName").GetString());
        Assert.Equal(element, read.RootElement.GetProperty("element").GetString());
        WithChart(sheet, chartName, chart => InspectNativeFont(chart, element, font =>
        {
            Assert.Equal("Arial", (object)font.Name);
            Assert.Equal(14.5, (object)font.Size);
            Assert.Equal(true, (object)font.Bold);
            Assert.Equal(true, (object)font.Italic);
            Assert.Equal(0xAC6824, Convert.ToInt32((object)font.Color, System.Globalization.CultureInfo.InvariantCulture));
        }));
        using var updated = JsonDocument.Parse(_fixture.Send("chartconfig.set-text-format", new
        {
            chartName,
            element,
            axis,
            bold = false,
            italic = false
        }).Result!);
        foreach (var format in updated.RootElement.GetProperty("formats").EnumerateArray())
        {
            Assert.Equal("Arial", format.GetProperty("fontName").GetString());
            Assert.Equal(14.5, format.GetProperty("fontSize").GetDouble());
            Assert.False(format.GetProperty("bold").GetBoolean());
            Assert.False(format.GetProperty("italic").GetBoolean());
            Assert.Equal("#2468AC", format.GetProperty("color").GetString());
        }
        WithChart(sheet, chartName, chart =>
        {
            Assert.True(chart.HasTitle);
            Assert.True(chart.HasLegend);
            Excel.ChartTitle? title = null;
            try
            {
                title = chart.ChartTitle;
                Assert.Equal("Revenue", title.Text);
            }
            finally { ComUtilities.Release(ref title); }
        });
    }

    [Fact]
    public void ReadTextFormat_ReportsMixedTitlePropertiesIncludingBlack()
    {
        var (sheet, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-title", new { chartName, title = "Revenue" });
        WithChart(sheet, chartName, chart =>
        {
            Excel.ChartTitle? title = null;
            Excel.Characters? first = null;
            Excel.Characters? rest = null;
            Excel.Font? firstFont = null;
            Excel.Font? restFont = null;
            try
            {
                title = chart.ChartTitle;
                first = title.Characters[1, 1];
                rest = title.Characters[2, 6];
                firstFont = first.Font;
                restFont = rest.Font;
                firstFont.Name = "Arial";
                firstFont.Size = 12;
                firstFont.Bold = true;
                firstFont.Italic = false;
                firstFont.Color = 0;
                restFont.Name = "Courier New";
                restFont.Size = 18;
                restFont.Bold = false;
                restFont.Italic = true;
                restFont.Color = 255;
            }
            finally
            {
                ComUtilities.Release(ref restFont);
                ComUtilities.Release(ref firstFont);
                ComUtilities.Release(ref rest);
                ComUtilities.Release(ref first);
                ComUtilities.Release(ref title);
            }
        });
        using var read = JsonDocument.Parse(_fixture.Send("chartconfig.get-text-format", new { chartName, element = "Title" }).Result!);
        var format = Assert.Single(read.RootElement.GetProperty("formats").EnumerateArray());
        Assert.Equal(["bold", "color", "fontName", "fontSize", "italic"],
            format.GetProperty("mixedFields").EnumerateArray().Select(v => v.GetString()).Order(StringComparer.Ordinal));
        foreach (var field in new[] { "fontName", "fontSize", "bold", "italic", "color" })
            Assert.True(!format.TryGetProperty(field, out var value) || value.ValueKind == JsonValueKind.Null);
    }

    [Fact]
    public void DataLabels_ReadsEachSeriesAndLeavesOtherSeriesUnchanged()
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-data-labels", new { chartName, showValue = true, labelPosition = "OutsideEnd" });
        _fixture.Send("chartconfig.set-text-format", new { chartName, element = "DataLabels", fontName = "Arial", fontSize = 12, color = "#000000" });
        _fixture.Send("chartconfig.set-text-format", new { chartName, element = "DataLabels", seriesIndex = 2, color = "#FF0000" });
        using var read = JsonDocument.Parse(_fixture.Send("chartconfig.get-text-format", new { chartName, element = "DataLabels", seriesIndex = 0 }).Result!);
        var formats = read.RootElement.GetProperty("formats");
        Assert.Equal(2, formats.GetArrayLength());
        Assert.Equal(1, formats[0].GetProperty("seriesIndex").GetInt32());
        Assert.Equal("#000000", formats[0].GetProperty("color").GetString());
        Assert.Equal(2, formats[1].GetProperty("seriesIndex").GetInt32());
        Assert.Equal("#FF0000", formats[1].GetProperty("color").GetString());
    }

    [Fact]
    public async Task AllSeries_MissingLabelsFailsBeforeChangingFirstSeries()
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-data-labels", new { chartName, showValue = true, seriesIndex = 1 });
        var before = _fixture.Send("chartconfig.get-text-format", new { chartName, element = "DataLabels", seriesIndex = 1 }).Result;
        var failure = await _fixture.SendForFailureAsync("chartconfig.set-text-format", new { chartName, element = "DataLabels", color = "#FF0000" });
        Assert.False(failure.Success);
        Assert.Contains("Series 2 has no data labels", failure.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(before, _fixture.Send("chartconfig.get-text-format", new { chartName, element = "DataLabels", seriesIndex = 1 }).Result);
    }

    [Theory]
    [InlineData("Title", null, "no title")]
    [InlineData("Legend", null, "no legend")]
    [InlineData("AxisTitle", "Value", "no title")]
    [InlineData("DataLabels", null, "no data labels")]
    public async Task MissingElements_FailWithoutCreatingThem(string element, string? axis, string message)
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-title", new { chartName, title = "" });
        _fixture.Send("chartconfig.show-legend", new { chartName, visible = false });
        var before = _fixture.Send("chart.read", new { chartName }).Result;
        foreach (var action in new[] { "get-text-format", "set-text-format" })
        {
            var args = new Dictionary<string, object?> { ["chartName"] = chartName, ["element"] = element, ["axis"] = axis };
            if (action == "set-text-format") args["fontSize"] = 14;
            var failure = await _fixture.SendForFailureAsync($"chartconfig.{action}", args);
            Assert.False(failure.Success);
            Assert.Contains(message, failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        }
        Assert.Equal(before, _fixture.Send("chart.read", new { chartName }).Result);
    }

    [Theory]
    [InlineData("""{"element":"Title"}""")]
    [InlineData("""{"element":"Title","fontName":""}""")]
    [InlineData("""{"element":"Title","fontSize":0}""")]
    [InlineData("""{"element":"Title","fontSize":410}""")]
    [InlineData("""{"element":"Title","color":"red"}""")]
    [InlineData("""{"element":"Title","color":"#12GG56"}""")]
    [InlineData("""{"element":"Title","color":"123"}""")]
    [InlineData("""{"element":"Title","bold":true,"seriesIndex":1}""")]
    [InlineData("""{"element":"Title","bold":true,"axis":"Value"}""")]
    [InlineData("""{"element":"AxisTitle","bold":true}""")]
    [InlineData("""{"element":"CategoryAxis","bold":true,"axis":"Value"}""")]
    [InlineData("""{"element":"ValueAxis","bold":true,"axis":"Category"}""")]
    [InlineData("""{"element":"DataLabels","bold":true,"seriesIndex":-1}""")]
    [InlineData("""{"element":"DataLabels","bold":true,"seriesIndex":3}""")]
    [InlineData("""{"element":"unknown","bold":true}""")]
    public async Task InvalidSettings_PreserveExistingChart(string json)
    {
        var (_, chartName) = CreateChart();
        var args = System.Text.Json.Nodes.JsonNode.Parse(json)!.AsObject();
        args["chartName"] = chartName;
        var before = _fixture.Send("chart.read", new { chartName }).Result;
        var failure = await _fixture.SendForFailureAsync("chartconfig.set-text-format", args);
        Assert.False(failure.Success);
        Assert.False(string.IsNullOrWhiteSpace(failure.ErrorMessage));
        Assert.Equal(before, _fixture.Send("chart.read", new { chartName }).Result);
    }

    [Fact]
    public async Task SaveReopen_PreservesTextAndSecondaryAxisSelection()
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-series-axis-group", new { chartName, seriesIndex = 2, axisGroup = "Secondary" });
        _fixture.Send("chartconfig.set-axis-title", new { chartName, axis = "ValueSecondary", title = "Secondary" });
        _fixture.Send("chartconfig.set-text-format", new { chartName, element = "ValueAxis", axis = "Value", color = "#123456" });
        _fixture.Send("chartconfig.set-text-format", new { chartName, element = "ValueAxis", axis = "ValueSecondary", color = "#ABCDEF" });
        _fixture.Send("chartconfig.set-text-format", new { chartName, element = "AxisTitle", axis = "ValueSecondary", fontName = "Arial", fontSize = 17, bold = true, italic = false, color = "#2468AC" });
        var primary = _fixture.Send("chartconfig.get-text-format", new { chartName, element = "ValueAxis" }).Result;
        var secondary = _fixture.Send("chartconfig.get-text-format", new { chartName, element = "ValueAxis", axis = "ValueSecondary" }).Result;
        using (var read = JsonDocument.Parse(primary!))
        {
            var format = Assert.Single(read.RootElement.GetProperty("formats").EnumerateArray());
            Assert.Equal("Value", format.GetProperty("axis").GetString());
            Assert.Equal("#123456", format.GetProperty("color").GetString());
        }
        using (var read = JsonDocument.Parse(secondary!))
        {
            var format = Assert.Single(read.RootElement.GetProperty("formats").EnumerateArray());
            Assert.Equal("ValueSecondary", format.GetProperty("axis").GetString());
            Assert.Equal("#ABCDEF", format.GetProperty("color").GetString());
        }
        var title = _fixture.Send("chartconfig.get-text-format", new { chartName, element = "AxisTitle", axis = "ValueSecondary" }).Result;
        await _fixture.SaveAndReopenAsync();
        Assert.Equal(primary, _fixture.Send("chartconfig.get-text-format", new { chartName, element = "ValueAxis" }).Result);
        Assert.Equal(secondary, _fixture.Send("chartconfig.get-text-format", new { chartName, element = "ValueAxis", axis = "ValueSecondary" }).Result);
        Assert.Equal(title, _fixture.Send("chartconfig.get-text-format", new { chartName, element = "AxisTitle", axis = "ValueSecondary" }).Result);
    }

    [Fact]
    public void DataLabels_ReadsMixedColourAcrossExistingLabels()
    {
        var (sheet, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-data-labels", new { chartName, showValue = true, seriesIndex = 1 });
        WithChart(sheet, chartName, chart =>
        {
            Excel.SeriesCollection? series = null;
            Excel.Series? item = null;
            Excel.DataLabels? labels = null;
            Excel.DataLabel? label = null;
            Excel.Font? font = null;
            try
            {
                series = (Excel.SeriesCollection)chart.SeriesCollection();
                item = series.Item(1);
                labels = (Excel.DataLabels)item.DataLabels();
                for (int i = 1; i <= labels.Count; i++)
                {
                    label = labels.Item(i);
                    font = label.Font;
                    font.Color = i == 1 ? 0 : 255;
                    ComUtilities.Release(ref font);
                    ComUtilities.Release(ref label);
                }
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref label);
                ComUtilities.Release(ref labels);
                ComUtilities.Release(ref item);
                ComUtilities.Release(ref series);
            }
        });
        using var read = JsonDocument.Parse(_fixture.Send("chartconfig.get-text-format", new { chartName, element = "DataLabels", seriesIndex = 1 }).Result!);
        var format = Assert.Single(read.RootElement.GetProperty("formats").EnumerateArray());
        Assert.Equal(JsonValueKind.Null, format.GetProperty("color").ValueKind);
        Assert.Equal("color", Assert.Single(format.GetProperty("mixedFields").EnumerateArray()).GetString());
    }

    [Theory]
    [InlineData("CategoryAxis", null)]
    [InlineData("ValueAxis", null)]
    [InlineData("AxisTitle", "Value")]
    public async Task PieChart_RejectsMissingAxesWithoutChangingChart(string element, string? axis)
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-chart-type", new { chartName, chartType = "Pie" });
        var before = _fixture.Send("chart.read", new { chartName }).Result;
        var failure = await _fixture.SendForFailureAsync("chartconfig.set-text-format", new { chartName, element, axis, fontSize = 14 });
        Assert.False(failure.Success);
        Assert.Contains("axis", failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, _fixture.Send("chart.read", new { chartName }).Result);
    }

    [Fact]
    public async Task ProtectedChart_FailsAndPreservesTitleFormatting()
    {
        var (sheet, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-title", new { chartName, title = "Revenue" });
        var before = _fixture.Send("chartconfig.get-text-format", new { chartName, element = "Title" }).Result;
        _fixture.Send("worksheetstyle.set-protection", new { sheetName = sheet, isProtected = true });
        var failure = await Record.ExceptionAsync(async () =>
        {
            var response = await _fixture.SendForFailureAsync("chartconfig.set-text-format", new
            {
                chartName,
                element = "Title",
                fontName = "Arial",
                fontSize = 14
            });
            Assert.False(response.Success);
            Assert.False(string.IsNullOrWhiteSpace(response.HResult));
            Assert.Contains(chartName, response.ErrorMessage, StringComparison.Ordinal);
            Assert.Contains("fontName", response.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal(before, _fixture.Send("chartconfig.get-text-format", new { chartName, element = "Title" }).Result);
        });
        var cleanup = Record.Exception(() => _fixture.Send("worksheetstyle.set-protection", new { sheetName = sheet, isProtected = false }));
        if (cleanup is not null) failure = PersistentServiceCleanupFailures.Combine(failure, cleanup);
        PersistentServiceCleanupFailures.ThrowIfFailed(failure);
    }

    [Fact]
    public void LinkedTitle_FormattingPreservesCellFormulaAndReadsActualFont()
    {
        var (sheet, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-title", new { chartName, title = "Revenue" });
        var formula = $"={sheet}!$A$2";
        WithChart(sheet, chartName, chart =>
        {
            Excel.ChartTitle? title = null;
            try { title = chart.ChartTitle; title.Formula = formula; }
            finally { ComUtilities.Release(ref title); }
        });
        using var read = JsonDocument.Parse(_fixture.Send("chartconfig.set-text-format", new
        {
            chartName,
            element = "Title",
            fontName = "Arial",
            fontSize = 14.5,
            bold = true,
            italic = true,
            color = "#2468AC"
        }).Result!);
        AssertFormats(read.RootElement, 1);
        WithChart(sheet, chartName, chart =>
        {
            Excel.ChartTitle? title = null;
            try
            {
                title = chart.ChartTitle;
                Assert.Equal(formula, title.Formula);
                Assert.Equal("A", title.Text);
            }
            finally { ComUtilities.Release(ref title); }
        });
    }

    private static void InspectNativeFont(Excel.Chart chart, string element, Action<Excel.Font> action)
    {
        Excel.ChartArea? area = null;
        Excel.ChartTitle? title = null;
        Excel.Axis? axis = null;
        Excel.AxisTitle? axisTitle = null;
        Excel.TickLabels? ticks = null;
        Excel.Legend? legend = null;
        Excel.SeriesCollection? series = null;
        Excel.Series? item = null;
        Excel.DataLabels? labels = null;
        Excel.Font? font = null;
        try
        {
            switch (element)
            {
                case "ChartArea": area = chart.ChartArea; font = area.Font; break;
                case "Title": title = chart.ChartTitle; font = title.Font; break;
                case "Legend": legend = chart.Legend; font = legend.Font; break;
                case "DataLabels":
                    series = (Excel.SeriesCollection)chart.SeriesCollection();
                    item = series.Item(1);
                    labels = (Excel.DataLabels)item.DataLabels();
                    Assert.True(labels.ShowValue);
                    font = labels.Font;
                    break;
                default:
                    axis = (Excel.Axis)chart.Axes(element == "CategoryAxis" ? Excel.XlAxisType.xlCategory : Excel.XlAxisType.xlValue);
                    if (element == "AxisTitle") { axisTitle = axis.AxisTitle; font = axisTitle.Font; Assert.Equal("Amount", axisTitle.Text); }
                    else { ticks = axis.TickLabels; font = ticks.Font; }
                    break;
            }
            action(font);
        }
        finally
        {
            ComUtilities.Release(ref font);
            ComUtilities.Release(ref labels);
            ComUtilities.Release(ref item);
            ComUtilities.Release(ref series);
            ComUtilities.Release(ref legend);
            ComUtilities.Release(ref ticks);
            ComUtilities.Release(ref axisTitle);
            ComUtilities.Release(ref axis);
            ComUtilities.Release(ref title);
            ComUtilities.Release(ref area);
        }
    }

    private static void AssertFormats(JsonElement root, int count)
    {
        Assert.True(root.GetProperty("success").GetBoolean());
        var formats = root.GetProperty("formats").EnumerateArray().ToArray();
        Assert.Equal(count, formats.Length);
        foreach (var format in formats)
        {
            Assert.Equal("Arial", format.GetProperty("fontName").GetString());
            Assert.Equal(14.5, format.GetProperty("fontSize").GetDouble());
            Assert.True(format.GetProperty("bold").GetBoolean());
            Assert.True(format.GetProperty("italic").GetBoolean());
            Assert.Equal("#2468AC", format.GetProperty("color").GetString());
            Assert.Empty(format.GetProperty("mixedFields").EnumerateArray());
        }
    }

    private (string Sheet, string Chart) CreateChart()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Text_{Guid.NewGuid():N}";
        _fixture.Send("range.set-values", new
        {
            sheetName = sheet,
            rangeAddress = "A1:C4",
            values = new object[][] { ["Category", "First", "Second"], ["A", 10, 100], ["B", 20, 200], ["C", 30, 300] }
        });
        _fixture.Send("chart.create-from-range", new { sheetName = sheet, sourceRangeAddress = "A1:C4", chartType = "ColumnClustered", chartName = name });
        return (sheet, name);
    }

    private void WithChart(string sheetName, string name, Action<Excel.Chart> action)
    {
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ChartObjects? objects = null;
            Excel.ChartObject? item = null;
            Excel.Chart? chart = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                objects = (Excel.ChartObjects)sheet.ChartObjects();
                item = objects.Item(name);
                chart = item.Chart;
                action(chart);
            }
            finally
            {
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref item);
                ComUtilities.Release(ref objects);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }
}
