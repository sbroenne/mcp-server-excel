using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "false")]
public sealed class WorksheetCommentToolTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task WorksheetStyle_SetAndClearComment_RoundsTripThroughMcp()
    {
        const string sessionId = "recording-session";
        var setCall = await CallAsync(
            "set-comment",
            sessionId,
            new()
            {
                ["text"] = "Quarterly update"
            },
            """{"success":true}""",
            "worksheetstyle.set-comment",
            """{"sheetName":"CommentSheet","cellAddress":"A1","text":"Quarterly update"}""");
        using (var args = RecordingToolTest.ParseArgs(
            setCall.Request,
            "worksheetstyle.set-comment",
            sessionId))
        {
            Assert.Equal("CommentSheet", args.RootElement.GetProperty("sheetName").GetString());
            Assert.Equal("A1", args.RootElement.GetProperty("cellAddress").GetString());
            Assert.Equal("Quarterly update", args.RootElement.GetProperty("text").GetString());
        }

        var getCall = await CallAsync(
            "get-comment",
            sessionId,
            [],
            """{"success":true,"hasComment":true,"text":"Quarterly update"}""",
            "worksheetstyle.get-comment",
            """{"sheetName":"CommentSheet","cellAddress":"A1"}""");
        using (var result = JsonDocument.Parse(getCall.JsonResult))
        {
            Assert.True(result.RootElement.GetProperty("hasComment").GetBoolean());
            Assert.Equal(
                "Quarterly update",
                result.RootElement.GetProperty("text").GetString());
        }

        var clearCall = await CallAsync(
            "clear-comment",
            sessionId,
            [],
            """{"success":true}""",
            "worksheetstyle.clear-comment",
            """{"sheetName":"CommentSheet","cellAddress":"A1"}""");
        Assert.Equal("worksheetstyle.clear-comment", clearCall.Request.Command);
    }

    private Task<RecordingProgramTransportFixture.CapturedToolCall> CallAsync(
        string action,
        string sessionId,
        Dictionary<string, object?> extraArguments,
        string result,
        string expectedCommand,
        string expectedArgsJson)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["workbook_session_id"] = sessionId,
            ["sheet_name"] = "CommentSheet",
            ["cell_address"] = "A1"
        };
        foreach (var (key, value) in extraArguments)
        {
            arguments.Add(key, value);
        }

        return _fixture.CallToolAsync(
            action == "get-comment" ? "worksheet_style_read" : "worksheet_style",
            arguments,
            RecordingToolTest.Success(result),
            expectedCommand,
            expectedArgsJson);
    }
}
