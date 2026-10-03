using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Timelines")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class TimelineProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Selection_PreservesInclusiveCalendarDates()
    {
        var selection = new TimelineSelectionOptions { StartDate = new DateTime(2024, 2, 1), EndDate = new DateTime(2024, 2, 29) };
        var call = await fixture.CallToolAsync("slicer", new Dictionary<string, object?>
        {
            ["action"] = "set-timeline-selection",
            ["session_id"] = "session-1",
            ["slicer_name"] = "Dates",
            ["timeline_selection"] = new { startDate = "2024-02-01", endDate = "2024-02-29" }
        }, RecordingToolTest.Success("""{"success":true}"""), "slicer.set-timeline-selection",
            JsonSerializer.Serialize(new { slicerName = "Dates", timelineSelection = selection }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Layout_PreservesExplicitFalseAndTimelineEnum()
    {
        var options = new SlicerUpdateOptions { Width = 350, Granularity = TimelineGranularity.Days, ShowHeader = false };
        var call = await fixture.CallToolAsync("slicer", new Dictionary<string, object?>
        {
            ["action"] = "update-slicer",
            ["session_id"] = "session-1",
            ["slicer_name"] = "Dates",
            ["slicer_options"] = new { width = 350, granularity = "Days", showHeader = false }
        }, RecordingToolTest.Success("""{"success":true}"""), "slicer.update-slicer",
            JsonSerializer.Serialize(new { slicerName = "Dates", slicerOptions = options }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_AdvertisesDistinctTypedControlPayloads()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "slicer");
        var properties = tool.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("timeline_selection", out _));
        Assert.True(properties.TryGetProperty("slicer_options", out _));
    }
}
