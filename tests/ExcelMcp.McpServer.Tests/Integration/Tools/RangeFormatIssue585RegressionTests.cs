using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Range;
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
                ["action"] = "format",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Formatting",
                ["range_addresses"] = (string[])["A1:J1"],
                ["format_options"] = new { bold = true, fontColor = "#FFFFFF", fillColor = "#1F4E79", fontName = (string?)null, orientation = (int?)null }
            },
            RecordingToolTest.Success("""{"success":true}"""),
            "rangeformat.format",
            JsonSerializer.Serialize(new
            {
                sheetName = "Formatting",
                rangeAddresses = (string[])["A1:J1"],
                formatOptions = new CellFormatOptions { Bold = true, FontColor = "#FFFFFF", FillColor = "#1F4E79" }
            }, ServiceProtocol.JsonOptions));

        using var args = RecordingToolTest.ParseArgs(
            call.Request,
            "rangeformat.format",
            sessionId);
        var root = args.RootElement;
        Assert.Equal("Formatting", root.GetProperty("sheetName").GetString());
        Assert.Equal("A1:J1", root.GetProperty("rangeAddresses")[0].GetString());
        var options = root.GetProperty("formatOptions");
        Assert.True(options.GetProperty("bold").GetBoolean());
        Assert.Equal("#FFFFFF", options.GetProperty("fontColor").GetString());
        Assert.Equal("#1F4E79", options.GetProperty("fillColor").GetString());
    }

    [Fact]
    public async Task FormatRange_InvalidColor_ReturnsTransparentFailureEnvelope()
    {
        const string sessionId = "recording-session";
        var call = await _fixture.CallToolAsync(
            "range_format",
            new Dictionary<string, object?>
            {
                ["action"] = "format",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Formatting",
                ["range_addresses"] = (string[])["A1:J1"],
                ["format_options"] = new { fillColor = "not-a-color" }
            },
            new ServiceResponse
            {
                Success = false,
                Command = "rangeformat.format",
                SessionId = sessionId,
                ErrorMessage = "Invalid color format: not-a-color",
                ErrorCategory = "InvalidInput",
                ExceptionType = "ArgumentException"
            },
            "rangeformat.format",
            JsonSerializer.Serialize(new
            {
                sheetName = "Formatting",
                rangeAddresses = (string[])["A1:J1"],
                formatOptions = new CellFormatOptions { FillColor = "not-a-color" }
            }, ServiceProtocol.JsonOptions));

        using var result = JsonDocument.Parse(call.JsonResult);
        var root = result.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal("ArgumentException", root.GetProperty("exceptionType").GetString());
        Assert.Equal("InvalidInput", root.GetProperty("errorCategory").GetString());
        Assert.Equal("rangeformat.format", root.GetProperty("command").GetString());
        Assert.Equal(sessionId, root.GetProperty("session_id").GetString());
        Assert.Contains(
            "not-a-color",
            root.GetProperty("errorMessage").GetString(),
            StringComparison.Ordinal);
    }
}
