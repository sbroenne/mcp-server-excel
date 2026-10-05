using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "false")]
public sealed class WorksheetProtectionToolTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task WorksheetStyle_SetProtection_RoundsTripsThroughMcp()
    {
        const string sessionId = "recording-session";
        var protectCall = await CallSetAsync(sessionId, true);
        using (var args = RecordingToolTest.ParseArgs(
            protectCall.Request,
            "sheet.set-protection",
            sessionId))
        {
            Assert.Equal("ProtectedSheet", args.RootElement.GetProperty("sheetName").GetString());
            Assert.True(args.RootElement.GetProperty("isProtected").GetBoolean());
        }

        var getCall = await _fixture.CallToolAsync(
            "worksheet_style_read",
            new Dictionary<string, object?>
            {
                ["action"] = "get-protection",
                ["session_id"] = sessionId,
                ["sheet_name"] = "ProtectedSheet"
            },
            RecordingToolTest.Success(
                """{"success":true,"isProtected":true}"""),
            "sheet.get-protection",
            """{"sheetName":"ProtectedSheet"}""");

        using (var args = RecordingToolTest.ParseArgs(
            getCall.Request,
            "sheet.get-protection",
            sessionId))
        {
            Assert.Equal("ProtectedSheet", args.RootElement.GetProperty("sheetName").GetString());
        }
        using (var result = JsonDocument.Parse(getCall.JsonResult))
        {
            Assert.True(result.RootElement.GetProperty("isProtected").GetBoolean());
        }

        var unprotectCall = await CallSetAsync(sessionId, false);
        using var unprotectArgs = RecordingToolTest.ParseArgs(
            unprotectCall.Request,
            "sheet.set-protection",
            sessionId);
        Assert.False(unprotectArgs.RootElement.GetProperty("isProtected").GetBoolean());
    }

    private Task<RecordingProgramTransportFixture.CapturedToolCall> CallSetAsync(
        string sessionId,
        bool isProtected) =>
        _fixture.CallToolAsync(
            "worksheet_style",
            new Dictionary<string, object?>
            {
                ["action"] = "set-protection",
                ["session_id"] = sessionId,
                ["sheet_name"] = "ProtectedSheet",
                ["is_protected"] = isProtected
            },
            RecordingToolTest.Success("""{"success":true}"""),
            "sheet.set-protection",
            isProtected
                ? """{"sheetName":"ProtectedSheet","isProtected":true}"""
                : """{"sheetName":"ProtectedSheet","isProtected":false}""");
}
