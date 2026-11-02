using System.Text.Json;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
[Trait("RequiresExcel", "false")]
public sealed class BatchToolProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Discovery_AdvertisesDefaultsAndNestedEntryNames()
    {
        var tool = Assert.Single(await fixture.ListToolsAsync(), tool => tool.Name == "batch");
        Assert.True(tool.ProtocolTool.Annotations!.DestructiveHint);
        Assert.False(tool.ProtocolTool.Annotations.ReadOnlyHint == true);
        Assert.False(tool.ProtocolTool.Annotations.IdempotentHint == true);
        var properties = tool.JsonSchema.GetProperty("properties");
        Assert.False(properties.GetProperty("stop_on_error").GetProperty("default").GetBoolean());
        Assert.Equal("array", properties.GetProperty("commands").GetProperty("type").GetString());
        var entry = properties.GetProperty("commands").GetProperty("items").GetProperty("properties");
        Assert.True(entry.TryGetProperty("command", out _));
        Assert.True(entry.TryGetProperty("sessionId", out _));
        Assert.True(entry.TryGetProperty("args", out _));
        Assert.Contains("stop_on_error:true", tool.Description);
    }

    [Theory]
    [InlineData("sheet.list", "worksheet_read", "list", true)]
    [InlineData("sheet.copy-to-file", "worksheet", "copy-to-file", false)]
    [InlineData("session.open", "file", "open", false)]
    [InlineData("session.save", "file", "save", true)]
    [InlineData("session.list", "file_read", "list", false)]
    [InlineData("session.notices", "file_read", "notices", true)]
    public void CommandIdentity_UsesTheAdvertisedToolAndActionSessionRequirement(
        string command, string tool, string action, bool requiresSession)
    {
        Assert.Equal((tool, action, requiresSession), McpActionContract.GetOperation(command));
    }

    [Fact]
    public async Task Run_ForwardsCommandAndReturnsIndexedResult()
    {
        var call = await fixture.CallToolAsync("batch", new()
        {
            ["action"] = "run",
            ["workbook_session_id"] = "recording-session",
            ["commands"] = new[]
            {
                new { command = "range.get-values", args = new { sheetName = "Data", rangeAddress = "A1" } }
            }
        }, RecordingToolTest.Success("""{"success":true,"values":[[42]]}"""),
            "range.get-values", """{"sheetName":"Data","rangeAddress":"A1"}""");

        using var output = JsonDocument.Parse(call.JsonResult);
        var root = output.RootElement;
        Assert.False(call.Result.IsError == true);
        Assert.True(JsonElement.DeepEquals(root, call.Result.StructuredContent!.Value));
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal(1, root.GetProperty("totalCount").GetInt32());
        Assert.Equal(1, root.GetProperty("executedCount").GetInt32());
        Assert.False(root.GetProperty("stoppedEarly").GetBoolean());
        var item = Assert.Single(root.GetProperty("results").EnumerateArray());
        Assert.Equal(0, item.GetProperty("index").GetInt32());
        Assert.Equal("range.get-values", item.GetProperty("command").GetString());
        Assert.True(item.GetProperty("success").GetBoolean());
        Assert.Equal("[[42]]", item.GetProperty("values").GetRawText());
    }

    [Theory]
    [InlineData(false, 3)]
    [InlineData(true, 2)]
    public async Task Failure_PreservesEarlierResultsAndOptionalStop(bool stopOnError, int expectedCount)
    {
        var arguments = Arguments(
            """
               [{"command":"range.get-values","args":{"sheetName":"Data","rangeAddress":"A1"}},
                {"command":"range.get-values","args":{"sheetName":"Data","rangeAddress":"B1"}},
                {"command":"range.get-values","args":{"sheetName":"Data","rangeAddress":"C1"}}]
            """,
            stopOnError);
        if (!stopOnError)
            arguments.Remove("stop_on_error");
        var call = await fixture.CallSequenceAsync("batch", arguments,
            RecordingToolTest.Success("""{"success":true,"values":[[17]]}"""),
            new ServiceResponse
            {
                Success = false,
                ErrorMessage = "Original COM failure",
                ErrorCategory = "ComInterop",
                ExceptionType = "COMException",
                HResult = "0x800A03EC",
                InnerError = "Original cause",
                PartialState = """{"retained":true}"""
            },
            RecordingToolTest.Success("""{"success":true,"values":[[false,0,null]]}"""));

        using var output = Parse(call.Result);
        var root = output.RootElement;
        Assert.True(call.Result.IsError);
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal(3, root.GetProperty("totalCount").GetInt32());
        Assert.Equal(expectedCount, root.GetProperty("executedCount").GetInt32());
        Assert.Equal(stopOnError, root.GetProperty("stoppedEarly").GetBoolean());
        Assert.Equal(expectedCount, call.Requests.Length);
        for (var index = 0; index < expectedCount; index++)
            RecordingToolTest.AssertRequest(call.Requests[index], "range.get-values", "recording-session",
                JsonSerializer.Serialize(new { sheetName = "Data", rangeAddress = $"{(char)('A' + index)}1" }));
        var results = root.GetProperty("results");
        Assert.Equal("[[17]]", results[0].GetProperty("values").GetRawText());
        var failure = results[1];
        Assert.False(failure.GetProperty("success").GetBoolean());
        Assert.Equal(1, failure.GetProperty("index").GetInt32());
        Assert.Equal("Original COM failure", failure.GetProperty("errorMessage").GetString());
        Assert.Equal("ComInterop", failure.GetProperty("errorCategory").GetString());
        Assert.Equal("0x800A03EC", failure.GetProperty("hresult").GetString());
        Assert.Equal("Original cause", failure.GetProperty("innerError").GetString());
        Assert.True(failure.GetProperty("partialState").GetProperty("retained").GetBoolean());
        if (!stopOnError)
        {
            Assert.True(results[2].GetProperty("success").GetBoolean());
            Assert.Equal(2, results[2].GetProperty("index").GetInt32());
            Assert.Equal("[[false,0,null]]", results[2].GetProperty("values").GetRawText());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DeliveredNegativeResult_FailsBatch(bool stopOnError)
    {
        var call = await fixture.CallSequenceAsync("batch", Arguments(
            """[{"command":"sheet.list"},{"command":"sheet.list"}]""", stopOnError),
            RecordingToolTest.Success("""{"success":false,"errorMessage":"Operation failed","errorCategory":"Conflict"}"""),
            RecordingToolTest.Success("""{"success":true,"sheets":[]}"""));
        using var output = Parse(call.Result);
        Assert.True(call.Result.IsError);
        Assert.Equal(stopOnError ? 1 : 2, call.Requests.Length);
        Assert.All(call.Requests, request =>
            RecordingToolTest.AssertRequest(request, "sheet.list", "recording-session", null));
        Assert.Equal("Conflict", output.RootElement.GetProperty("results")[0].GetProperty("errorCategory").GetString());
    }

    [Theory]
    [InlineData("session.open")]
    [InlineData("session.create")]
    [InlineData("session.attach")]
    public async Task SessionStartup_CapturesDefaultAndHonorsEntryOverride(string command)
    {
        var arguments = Arguments(JsonSerializer.Serialize(new object[]
        {
            new
            {
                command,
                args = command == "session.attach"
                    ? JsonSerializer.SerializeToElement(new { workbookId = "discovered-workbook" })
                    : JsonSerializer.SerializeToElement(new { filePath = @"C:\fixture.xlsx" })
            },
            new { command = "sheet.list", sessionId = "override-session" },
            new { command = "sheet.list" }
        }));
        arguments.Remove("workbook_session_id");
        var call = await fixture.CallSequenceAsync("batch", arguments,
            RecordingToolTest.Success("""{"success":true,"sessionId":"new-session"}"""),
            RecordingToolTest.Success("""{"success":true,"sheets":[]}"""),
            RecordingToolTest.Success("""{"success":true,"sheets":[]}"""));
        using var output = Parse(call.Result);
        Assert.False(call.Result.IsError == true);
        Assert.Equal(3, call.Requests.Length);
        var commands = Assert.IsType<JsonElement>(arguments["commands"]);
        RecordingToolTest.AssertRequest(call.Requests[0], command, null, commands[0].GetProperty("args").GetRawText());
        RecordingToolTest.AssertRequest(call.Requests[1], "sheet.list", "override-session", null);
        RecordingToolTest.AssertRequest(call.Requests[2], "sheet.list", "new-session", null);
        Assert.Equal("new-session", output.RootElement.GetProperty("results")[0].GetProperty("sessionId").GetString());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Close_ClearsOnlySuccessfullyClosedDefaultSession(bool success)
    {
        var call = await fixture.CallSequenceAsync("batch", Arguments(
            """[{"command":"session.close"},{"command":"sheet.list"}]"""),
            new ServiceResponse { Success = success, Result = JsonSerializer.Serialize(new { success }) },
            RecordingToolTest.Success("""{"success":true,"sheets":[]}"""));
        using var output = Parse(call.Result);
        Assert.Equal(2, call.Requests.Length);
        RecordingToolTest.AssertRequest(call.Requests[0], "session.close", "recording-session", null);
        RecordingToolTest.AssertRequest(call.Requests[1], "sheet.list", success ? null : "recording-session", null);
        Assert.Equal(!success, call.Result.IsError == true);
    }

    [Fact]
    public async Task ExplicitDefault_IsNotChangedByOpeningOrClosingAnotherSession()
    {
        var call = await fixture.CallSequenceAsync("batch", Arguments(
            """
            [{"command":"session.create","args":{"filePath":"C:\\fixture.xlsx"}},
             {"command":"session.close","sessionId":"another-session"},
             {"command":"sheet.list"}]
            """),
            RecordingToolTest.Success("""{"success":true,"sessionId":"another-session"}"""),
            RecordingToolTest.Success("""{"success":true}"""),
            RecordingToolTest.Success("""{"success":true,"sheets":[]}"""));
        using var output = Parse(call.Result);
        Assert.False(call.Result.IsError == true);
        Assert.Equal(3, call.Requests.Length);
        RecordingToolTest.AssertRequest(call.Requests[0], "session.create", null,
            """{"filePath":"C:\\fixture.xlsx"}""");
        RecordingToolTest.AssertRequest(call.Requests[1], "session.close", "another-session", null);
        RecordingToolTest.AssertRequest(call.Requests[2], "sheet.list", "recording-session", null);
    }

    [Theory]
    [InlineData("""[{"command":"sheet.list","args":"wrong"}]""", "args")]
    [InlineData("""[{"command":"sheet.list","sessionId":" "}]""", "sessionId")]
    [InlineData("""[{"command":"batch.run"}]""", "batch")]
    [InlineData("""[{"command":"service.shutdown"}]""", "service.shutdown")]
    [InlineData("""[{"command":"service.status"}]""", "service.status")]
    [InlineData("""[{"command":"service.ping"}]""", "service.ping")]
    [InlineData("""[{"command":"diag.echo","args":{"message":"hello"}}]""", "diag.echo")]
    public async Task InvalidEntry_ReportsInputErrorWithoutDispatch(string commands, string expectedMessage)
    {
        var result = await fixture.CallResultWithoutDispatchAsync("batch", Arguments(commands));
        using var output = Parse(result);
        Assert.True(result.IsError);
        var item = Assert.Single(output.RootElement.GetProperty("results").EnumerateArray());
        Assert.Equal("InvalidInput", item.GetProperty("errorCategory").GetString());
        Assert.Contains(expectedMessage, item.GetProperty("errorMessage").GetString());
    }

    [Fact]
    public async Task Failure_ExplainsNoRollbackWithoutPromisingRetainedEdits()
    {
        var call = await fixture.CallSequenceAsync("batch", Arguments(
            """[{"command":"session.close"},{"command":"sheet.list"}]"""),
            RecordingToolTest.Success("""{"success":true}"""),
            new ServiceResponse { Success = false, ErrorCategory = "SessionNotFound", ErrorMessage = "No active session" });
        using var output = Parse(call.Result);
        var message = output.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("does not roll back", message, StringComparison.Ordinal);
        Assert.DoesNotContain("earlier changes remain", message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ValidationFailure_IsIndexedAndDoesNotDispatchInvalidEntry(bool stopOnError)
    {
        var call = await fixture.CallSequenceAsync("batch", Arguments(
            """
               [{"command":"range.get-values","args":{"sheet":"Wrong","rangeAddress":"A1"}},
                {"command":"sheet.list"}]
            """, stopOnError),
            RecordingToolTest.Success("""{"success":true,"sheets":[]}"""));
        using var output = Parse(call.Result);
        Assert.True(call.Result.IsError);
        Assert.Equal(stopOnError ? 0 : 1, call.Requests.Length);
        if (!stopOnError)
            RecordingToolTest.AssertRequest(Assert.Single(call.Requests), "sheet.list", "recording-session", null);
        var rejected = output.RootElement.GetProperty("results")[0];
        Assert.Equal(0, rejected.GetProperty("index").GetInt32());
        Assert.Equal("InvalidInput", rejected.GetProperty("errorCategory").GetString());
        Assert.Contains("sheet", rejected.GetProperty("errorMessage").GetString());
    }

    [Theory]
    [InlineData("[]")]
    [InlineData("[null]")]
    [InlineData("[{}]")]
    [InlineData("""[{"command":"sheet.list"},{}]""")]
    public async Task MalformedList_IsRejectedBeforeAnyDispatch(string commands)
    {
        var result = await fixture.CallResultWithoutDispatchAsync("batch", Arguments(commands));
        using var output = Parse(result);
        Assert.True(result.IsError);
        Assert.Equal("InvalidInput", output.RootElement.GetProperty("errorCategory").GetString());
    }

    [Fact]
    public async Task CancelledResponse_StopsEvenWithoutStopOnError()
    {
        var call = await fixture.CallSequenceAsync("batch", Arguments(
            """[{"command":"sheet.list"},{"command":"sheet.list"}]"""),
            new ServiceResponse { Success = false, ErrorCategory = "Cancelled", ErrorMessage = "Cancelled by owner" });
        using var output = Parse(call.Result);
        Assert.True(call.Result.IsError);
        RecordingToolTest.AssertRequest(Assert.Single(call.Requests), "sheet.list", "recording-session", null);
        Assert.True(output.RootElement.GetProperty("stoppedEarly").GetBoolean());
    }

    private static Dictionary<string, object?> Arguments(string commands, bool stopOnError = false) => new()
    {
        ["action"] = "run",
        ["workbook_session_id"] = "recording-session",
        ["stop_on_error"] = stopOnError,
        ["commands"] = JsonSerializer.Deserialize<JsonElement>(commands)
    };

    private static JsonDocument Parse(CallToolResult result)
    {
        var output = JsonDocument.Parse(Assert.Single(result.Content.OfType<TextContentBlock>()).Text);
        Assert.True(JsonElement.DeepEquals(output.RootElement, result.StructuredContent!.Value));
        return output;
    }
}
