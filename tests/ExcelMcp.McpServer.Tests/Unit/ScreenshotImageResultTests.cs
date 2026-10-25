using System.Text.Json;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Screenshot")]
[Trait("RequiresExcel", "false")]
public sealed class ScreenshotImageResultTests
{
    [Fact]
    public void CreateToolResult_Failure_HasStructuredProtocolError()
    {
        var result = ScreenshotImageResult.CreateToolResult("""{"success":false,"errorMessage":"Screenshot unavailable."}""");
        Assert.True(result.IsError);
        Assert.NotNull(result.StructuredContent);
        Assert.False(result.StructuredContent.Value.GetProperty("success").GetBoolean());
    }

    [Fact]
    public void CreateToolResult_MissingImage_RejectsMalformedSuccess()
    {
        Assert.Throws<InvalidOperationException>(() =>
            ScreenshotImageResult.CreateToolResult("""{"success":true}"""));
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

        var result = ScreenshotImageResult.CreateToolResult(json);

        Assert.NotEqual(true, result.IsError);
        Assert.Single(result.Content.OfType<ImageContentBlock>());
        var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
        Assert.Contains(truncationMessage, text, StringComparison.Ordinal);
        var structured = Assert.IsType<JsonElement>(result.StructuredContent);
        Assert.True(structured.GetProperty("success").GetBoolean());
        Assert.Equal("image/png", structured.GetProperty("mimeType").GetString());
        Assert.False(structured.TryGetProperty("imageBase64", out _));
    }
}
