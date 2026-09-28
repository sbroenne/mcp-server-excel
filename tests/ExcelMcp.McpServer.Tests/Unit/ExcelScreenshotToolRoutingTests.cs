using System.Text.Json;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Screenshot")]
[Trait("RequiresExcel", "false")]
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

    [Fact]
    public void CreateToolResult_SuccessWithTruncationMessage_ReturnsMessageToCaller()
    {
        const string truncationMessage =
            "The range was too large to capture in full and was truncated to its top-left portion";
        var screenshot = new ScreenshotResult
        {
            Success = true,
            ImageBase64 = Convert.ToBase64String([1, 2, 3]),
            MimeType = "image/png",
            Width = 120,
            Height = 80,
            SheetName = "Summary",
            RangeAddress = "$A$1:$BA$20",
            Message = truncationMessage
        };
        string json = JsonSerializer.Serialize(screenshot, ExcelToolsBase.JsonOptions);

        var result = ExcelScreenshotTool.CreateToolResult(json);

        Assert.NotEqual(true, result.IsError);
        Assert.Single(result.Content.OfType<ImageContentBlock>());
        var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
        Assert.Contains(truncationMessage, text, StringComparison.Ordinal);
    }
}
