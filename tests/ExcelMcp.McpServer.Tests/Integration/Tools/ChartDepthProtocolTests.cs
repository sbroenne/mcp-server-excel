using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "ChartDepth")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ChartDepthProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task ErrorBars_PreservesTypedOptionsAndNativeDirection()
    {
        var options = new ChartErrorBarOptions { Direction = ChartErrorBarDirection.X, Kind = ChartErrorBarKind.Fixed, Amount = 2, EndStyle = ChartErrorBarEndStyle.NoCap };
        var call = await fixture.CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-error-bars",
            ["session_id"] = "session-1",
            ["chart_name"] = "Sales",
            ["series_index"] = 2,
            ["error_bar_options"] = new { direction = "X", kind = "Fixed", amount = 2, endStyle = "NoCap" }
        }, RecordingToolTest.Success("""{"success":true}"""), "chartconfig.set-error-bars",
            JsonSerializer.Serialize(new { chartName = "Sales", seriesIndex = 2, errorBarOptions = options }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task PointFormat_PreservesIndicesAndTypedMarkerSettings()
    {
        var options = new ChartPointOptions { FillColor = "#FF0000", MarkerStyle = MarkerStyle.Diamond, MarkerSize = 14 };
        var call = await fixture.CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-point-format",
            ["session_id"] = "session-1",
            ["chart_name"] = "Sales",
            ["series_index"] = 2,
            ["point_index"] = 3,
            ["point_options"] = new { fillColor = "#FF0000", markerStyle = "Diamond", markerSize = 14 }
        }, RecordingToolTest.Success("""{"success":true}"""), "chartconfig.set-point-format",
            JsonSerializer.Serialize(new { chartName = "Sales", seriesIndex = 2, pointIndex = 3, pointOptions = options }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task ImageExport_PreservesFormatAndOverwrite()
    {
        var call = await fixture.CallToolAsync("chart", new Dictionary<string, object?>
        {
            ["action"] = "export-image",
            ["session_id"] = "session-1",
            ["chart_name"] = "Sales",
            ["target_path"] = "sales.jpg",
            ["image_format"] = "Jpeg",
            ["overwrite"] = true
        }, RecordingToolTest.Success("""{"success":true}"""), "chart.export-image",
            """{"chartName":"Sales","targetPath":"sales.jpg","imageFormat":"Jpeg","overwrite":true}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_AdvertisesDistinctTypedPayloadsAndReadLimits()
    {
        var tools = await fixture.ListToolsAsync();
        var config = Assert.Single(tools, tool => tool.Name == "chart_config");
        var properties = config.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("error_bar_options", out _));
        Assert.True(properties.TryGetProperty("point_options", out _));
        Assert.True(properties.TryGetProperty("axis_group", out _));
        Assert.Contains("settingsReadable", config.Description, StringComparison.Ordinal);
    }
}
