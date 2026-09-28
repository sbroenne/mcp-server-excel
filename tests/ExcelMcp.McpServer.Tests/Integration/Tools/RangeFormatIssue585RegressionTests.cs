using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangeFormatIssue585RegressionTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task FormatRange_ApplicableExplicitNulls_AreAcceptedViaMcpProtocol()
    {
        const string sessionId = "recording-session";
        var call = await _fixture.CallToolAsync(
            "range_format",
            new Dictionary<string, object?>
            {
                ["action"] = "format-range",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Formatting",
                ["range_address"] = "A1:J1",
                ["font_name"] = null,
                ["font_size"] = null,
                ["bold"] = true,
                ["italic"] = null,
                ["underline"] = null,
                ["font_color"] = "#FFFFFF",
                ["fill_color"] = "#1F4E79",
                ["border_style"] = null,
                ["border_color"] = null,
                ["border_weight"] = null,
                ["horizontal_alignment"] = null,
                ["vertical_alignment"] = null,
                ["wrap_text"] = null,
                ["orientation"] = null
            },
            RecordingToolTest.Success("""{"success":true}"""),
            "rangeformat.format-range",
            """{"sheetName":"Formatting","rangeAddress":"A1:J1","bold":true,"fontColor":"#FFFFFF","fillColor":"#1F4E79"}""");

        using var args = RecordingToolTest.ParseArgs(
            call.Request,
            "rangeformat.format-range",
            sessionId);
        var root = args.RootElement;
        Assert.Equal("Formatting", root.GetProperty("sheetName").GetString());
        Assert.Equal("A1:J1", root.GetProperty("rangeAddress").GetString());
        Assert.True(root.GetProperty("bold").GetBoolean());
        Assert.Equal("#FFFFFF", root.GetProperty("fontColor").GetString());
        Assert.Equal("#1F4E79", root.GetProperty("fillColor").GetString());
        Assert.False(root.TryGetProperty("fontName", out _));
        Assert.False(root.TryGetProperty("orientation", out _));
    }

    [Fact]
    public async Task FormatRange_InvalidColor_ReturnsTransparentFailureEnvelope()
    {
        const string sessionId = "recording-session";
        var call = await _fixture.CallToolAsync(
            "range_format",
            new Dictionary<string, object?>
            {
                ["action"] = "format-range",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Formatting",
                ["range_address"] = "A1:J1",
                ["fill_color"] = "not-a-color"
            },
            new ServiceResponse
            {
                Success = false,
                Command = "rangeformat.format-range",
                SessionId = sessionId,
                ErrorMessage = "Invalid color format: not-a-color",
                ErrorCategory = "InvalidInput",
                ExceptionType = "ArgumentException"
            },
            "rangeformat.format-range",
            """{"sheetName":"Formatting","rangeAddress":"A1:J1","fillColor":"not-a-color"}""");

        using var result = JsonDocument.Parse(call.JsonResult);
        var root = result.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal("ArgumentException", root.GetProperty("exceptionType").GetString());
        Assert.Equal("InvalidInput", root.GetProperty("errorCategory").GetString());
        Assert.Equal("rangeformat.format-range", root.GetProperty("command").GetString());
        Assert.Equal(sessionId, root.GetProperty("sessionId").GetString());
        Assert.Contains(
            "not-a-color",
            root.GetProperty("errorMessage").GetString(),
            StringComparison.Ordinal);
    }
}
