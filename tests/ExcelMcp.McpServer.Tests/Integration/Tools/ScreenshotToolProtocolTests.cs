using System.Text.Json;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Screenshot")]
[Trait("Feature", "GeneratedContracts")]
[Trait("RequiresExcel", "false")]
public sealed class ScreenshotToolProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Capture_KeepsNativeImageAndCompleteDiagnostics()
    {
        var call = await fixture.CallToolAsync("screenshot", new()
        {
            ["action"] = "capture",
            ["workbook_session_id"] = "session-1"
        }, RecordingToolTest.Success(
            """{"success":true,"imageBase64":"AQID","mimeType":"image/png","width":120,"height":80,"sheetName":"Summary","rangeAddress":"$A$1:$B$2","message":"Truncated to the top-left portion.","warnings":["Inspect the captured scope."],"partialState":{"captured":true}}"""),
            "screenshot.capture", "{}");

        Assert.False(call.Result.IsError == true);
        var image = Assert.Single(call.Result.Content.OfType<ImageContentBlock>());
        Assert.Equal("image/png", image.MimeType);
        Assert.Equal(new byte[] { 1, 2, 3 },
            Convert.FromBase64String(System.Text.Encoding.UTF8.GetString(image.Data.Span)));
        var metadata = Assert.IsType<JsonElement>(call.Result.StructuredContent);
        Assert.Equal("image/png", metadata.GetProperty("mimeType").GetString());
        Assert.Equal(120, metadata.GetProperty("width").GetInt32());
        Assert.Equal(80, metadata.GetProperty("height").GetInt32());
        Assert.Equal("$A$1:$B$2", metadata.GetProperty("rangeAddress").GetString());
        Assert.Equal("Truncated to the top-left portion.", metadata.GetProperty("message").GetString());
        Assert.Equal("Inspect the captured scope.", Assert.Single(metadata.GetProperty("warnings").EnumerateArray()).GetString());
        Assert.True(metadata.GetProperty("partialState").GetProperty("captured").GetBoolean());
        Assert.False(metadata.TryGetProperty("imageBase64", out _));
    }

    [Theory]
    [InlineData("capture", false)]
    [InlineData("capture-sheet", false)]
    [InlineData("capture", true)]
    [InlineData("capture-sheet", true)]
    public async Task Capture_ReturnsNativeImageAndPreservesActionInputs(string action, bool explicitInputs)
    {
        byte[] imageBytes = [1, 2, 3, 4];
        var screenshot = new ScreenshotResult
        {
            Success = true,
            ImageBase64 = Convert.ToBase64String(imageBytes),
            MimeType = "image/png",
            Width = 120,
            Height = 80,
            SheetName = "Summary",
            RangeAddress = "$B$2:$C$4",
            Message = "Truncated to the top-left portion."
        };
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1"
        };
        var expected = new Dictionary<string, object?>();
        if (explicitInputs)
        {
            arguments["sheet_name"] = expected["sheetName"] = "Summary";
            arguments["quality"] = expected["quality"] = "High";
            if (action == "capture")
                arguments["range_address"] = expected["rangeAddress"] = "B2:C4";
        }
        var call = await fixture.CallToolAsync("screenshot", arguments,
            RecordingToolTest.Success(JsonSerializer.Serialize(screenshot, ExcelToolsBase.JsonOptions)),
            $"screenshot.{action}", JsonSerializer.Serialize(expected));

        Assert.False(call.Result.IsError == true);
        var image = Assert.Single(call.Result.Content.OfType<ImageContentBlock>());
        Assert.Equal("image/png", image.MimeType);
        Assert.Equal(imageBytes, Convert.FromBase64String(System.Text.Encoding.UTF8.GetString(image.Data.Span)));
        var text = Assert.Single(call.Result.Content.OfType<TextContentBlock>()).Text;
        Assert.Contains(screenshot.Message, text, StringComparison.Ordinal);
        Assert.Contains("120x80px", text, StringComparison.Ordinal);
        var metadata = Assert.IsType<JsonElement>(call.Result.StructuredContent);
        Assert.True(metadata.GetProperty("success").GetBoolean());
        Assert.Equal("Summary", metadata.GetProperty("sheetName").GetString());
        Assert.Equal("$B$2:$C$4", metadata.GetProperty("rangeAddress").GetString());
        Assert.Equal(120, metadata.GetProperty("width").GetInt32());
        Assert.Equal(80, metadata.GetProperty("height").GetInt32());
        Assert.Equal("image/png", metadata.GetProperty("mimeType").GetString());
        Assert.Equal(screenshot.Message, metadata.GetProperty("message").GetString());
        Assert.False(metadata.TryGetProperty("imageBase64", out _));
    }

    [Fact]
    public async Task CaptureFailure_PreservesServiceErrorDetails()
    {
        var call = await fixture.CallToolAsync("screenshot", new()
        {
            ["action"] = "capture",
            ["workbook_session_id"] = "session-1"
        }, new ServiceResponse
        {
            Success = false,
            Command = "screenshot.capture",
            ErrorMessage = "Desktop unavailable.",
            ErrorCategory = "Prerequisite",
            ExceptionType = "InvalidOperationException",
            PartialState = """{"captured":false}"""
        }, "screenshot.capture", "{}");
        Assert.True(call.Result.IsError);
        Assert.Empty(call.Result.Content.OfType<ImageContentBlock>());
        var metadata = Assert.IsType<JsonElement>(call.Result.StructuredContent);
        Assert.Equal("Desktop unavailable.", metadata.GetProperty("errorMessage").GetString());
        Assert.Equal("Prerequisite", metadata.GetProperty("errorCategory").GetString());
        Assert.Equal("InvalidOperationException", metadata.GetProperty("exceptionType").GetString());
        Assert.False(metadata.GetProperty("partialState").GetProperty("captured").GetBoolean());
    }

    [Theory]
    [InlineData("capture-sheet", "range_address", "B2:C4")]
    [InlineData("capture", "quality", "invalid-quality")]
    public async Task InvalidInputs_AreRejectedWithoutDispatch(string action, string input, string value)
    {
        var result = await fixture.CallResultWithoutDispatchAsync("screenshot", new()
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1",
            [input] = value
        });
        Assert.True(result.IsError);
        Assert.Contains(input, Assert.Single(result.Content.OfType<TextContentBlock>()).Text,
            StringComparison.Ordinal);
    }
}
