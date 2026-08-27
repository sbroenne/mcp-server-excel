using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Screenshot")]
public sealed class ExcelScreenshotToolRoutingTests
{
    [Fact]
    public void RouteScreenshotAction_CaptureSheet_DoesNotSupplyRangeAddress()
    {
        string? command = null;
        object? arguments = null;

        var response = ExcelScreenshotTool.RouteScreenshotAction(
            ScreenshotAction.CaptureSheet,
            "session-1",
            "Summary",
            "A1:Z30",
            ScreenshotQuality.Medium,
            (routedCommand, _, routedArguments) =>
            {
                command = routedCommand;
                arguments = routedArguments;
                return """{"success":true}""";
            });

        Assert.Equal("""{"success":true}""", response);
        Assert.Equal("screenshot.capture-sheet", command);

        using var json = JsonDocument.Parse(JsonSerializer.Serialize(arguments, ExcelToolsBase.JsonOptions));
        Assert.False(json.RootElement.TryGetProperty("rangeAddress", out _));
    }

    [Fact]
    public void RouteScreenshotAction_Capture_SuppliesRangeAddress()
    {
        string? command = null;
        object? arguments = null;

        var response = ExcelScreenshotTool.RouteScreenshotAction(
            ScreenshotAction.CaptureRange,
            "session-1",
            "Summary",
            "B2:C4",
            ScreenshotQuality.High,
            (routedCommand, _, routedArguments) =>
            {
                command = routedCommand;
                arguments = routedArguments;
                return """{"success":true}""";
            });

        Assert.Equal("""{"success":true}""", response);
        Assert.Equal("screenshot.capture", command);

        using var json = JsonDocument.Parse(JsonSerializer.Serialize(arguments, ExcelToolsBase.JsonOptions));
        Assert.Equal("B2:C4", json.RootElement.GetProperty("rangeAddress").GetString());
    }
}
