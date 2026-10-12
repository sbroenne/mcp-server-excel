using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativeChartAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task NativeComboPointsErrorBarsAndImage_ExposeExactResults()
    {
        await ValuesAsync("AA1:AC4", """[["Category","First","Second"],["A",10,100],["B",20,200],["C",30,300]]""");
        var created = await CommandAsync("chart", "create-from-range", "--sheet-name", "Data", "--source-range-address", "AA1:AC4",
            "--chart-type", "ColumnClustered", "--chart-name", "DepthChart");
        Assert.Equal("DepthChart", Text(created, "chartName"));
        await ChartAsync("set-plot-options", "--plot-by", "Rows");
        await ChartAsync("set-series-chart-type", "--series-index", "2", "--chart-type", "LineMarkers");
        var axis = await ChartAsync("set-series-axis-group", "--series-index", "2", "--axis-group", "Secondary");
        Assert.Equal("Secondary", Text(axis, "axisGroup"));
        var settings = await ChartAsync("get-series-settings", "--series-index", "2");
        Assert.Equal("LineMarkers", Text(settings, "chartType"));
        Assert.Equal("Secondary", Text(settings, "axisGroup"));
        Assert.Equal(2, Number(settings, "pointCount"));
        Assert.Equal("B", Text(settings, "name"));
        var point = await ChartAsync("set-point-format", "--series-index", "1", "--point-index", "2",
            "--point-options", """{"fillColor":"#FF0000","lineColor":"#0000FF","lineWeight":2}""");
        Assert.Equal("#FF0000", Text(point, "fillColor"));
        var bars = await ChartAsync("set-error-bars", "--series-index", "1",
            "--error-bar-options", """{"kind":"Fixed","amount":2,"endStyle":"NoCap"}""");
        Assert.True(Bool(bars, "hasErrorBars"));
        Assert.Equal("NoCap", Text(bars, "endStyle"));
        var limits = await ChartAsync("get-error-bars", "--series-index", "1");
        Assert.True(Bool(limits, "hasErrorBars"));
        Assert.False(Bool(limits, "settingsReadable"));
        await CommandAsync("chart", "export-image", "--chart-name", "DepthChart", "--target-path", ChartImage);
        Assert.True(new FileInfo(ChartImage).Length > 1000);
    }

    private Task<System.Text.Json.JsonElement> ChartAsync(string action, params string[] arguments) =>
        CommandAsync("chartconfig", action, ["--chart-name", "DepthChart", .. arguments]);
}
