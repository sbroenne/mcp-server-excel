using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "false")]
public sealed class WorkbookToolTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task Workbook_SetProtection_RoundsTripsThroughMcp()
    {
        const string sessionId = "recording-session";
        var protectedCall = await _fixture.CallToolAsync(
            "workbook",
            new Dictionary<string, object?>
            {
                ["action"] = "set-protection",
                ["session_id"] = sessionId,
                ["is_protected"] = true
            },
            Success("""{"success":true,"isProtected":true}"""),
            "workbook.set-protection",
            """{"isProtected":true}""");

        AssertRequest(
            protectedCall.Request,
            "workbook.set-protection",
            sessionId,
            args => Assert.True(args.GetProperty("isProtected").GetBoolean()));
        using (var result = JsonDocument.Parse(protectedCall.JsonResult))
        {
            Assert.True(result.RootElement.GetProperty("isProtected").GetBoolean());
        }

        var getCall = await _fixture.CallToolAsync(
            "workbook",
            new Dictionary<string, object?>
            {
                ["action"] = "get-protection",
                ["session_id"] = sessionId
            },
            Success("""{"success":true,"isProtected":true}"""),
            "workbook.get-protection",
            null);

        AssertRequestWithoutArgs(
            getCall.Request,
            "workbook.get-protection",
            sessionId);
        using (var result = JsonDocument.Parse(getCall.JsonResult))
        {
            Assert.True(result.RootElement.GetProperty("isProtected").GetBoolean());
        }

        var unprotectedCall = await _fixture.CallToolAsync(
            "workbook",
            new Dictionary<string, object?>
            {
                ["action"] = "set-protection",
                ["session_id"] = sessionId,
                ["is_protected"] = false
            },
            Success("""{"success":true,"isProtected":false}"""),
            "workbook.set-protection",
            """{"isProtected":false}""");

        AssertRequest(
            unprotectedCall.Request,
            "workbook.set-protection",
            sessionId,
            args => Assert.False(args.GetProperty("isProtected").GetBoolean()));
    }

    [Fact]
    public async Task Workbook_SetViewOptions_RoundsTripsThroughMcp()
    {
        const string sessionId = "recording-session";
        var setCall = await _fixture.CallToolAsync(
            "workbook",
            new Dictionary<string, object?>
            {
                ["action"] = "set-view-options",
                ["session_id"] = sessionId,
                ["display_gridlines"] = false,
                ["display_headings"] = true
            },
            Success("""{"success":true}"""),
            "workbook.set-view-options",
            """{"displayGridlines":false,"displayHeadings":true}""");

        AssertRequest(
            setCall.Request,
            "workbook.set-view-options",
            sessionId,
            args =>
            {
                Assert.False(args.GetProperty("displayGridlines").GetBoolean());
                Assert.True(args.GetProperty("displayHeadings").GetBoolean());
            });

        var getCall = await _fixture.CallToolAsync(
            "workbook",
            new Dictionary<string, object?>
            {
                ["action"] = "get-view-options",
                ["session_id"] = sessionId
            },
            Success(
                """{"success":true,"displayGridlines":false,"displayHeadings":true}"""),
            "workbook.get-view-options",
            null);

        AssertRequestWithoutArgs(
            getCall.Request,
            "workbook.get-view-options",
            sessionId);
        using var result = JsonDocument.Parse(getCall.JsonResult);
        Assert.False(result.RootElement.GetProperty("displayGridlines").GetBoolean());
        Assert.True(result.RootElement.GetProperty("displayHeadings").GetBoolean());
    }

    private static ServiceResponse Success(string result) =>
        new()
        {
            Success = true,
            Result = result
        };

    private static void AssertRequest(
        ServiceRequest request,
        string command,
        string sessionId,
        Action<JsonElement> assertArgs)
    {
        Assert.Equal(command, request.Command);
        Assert.Equal(sessionId, request.SessionId);
        Assert.NotNull(request.Args);
        using var args = JsonDocument.Parse(request.Args);
        assertArgs(args.RootElement);
    }

    private static void AssertRequestWithoutArgs(
        ServiceRequest request,
        string command,
        string sessionId)
    {
        Assert.Equal(command, request.Command);
        Assert.Equal(sessionId, request.SessionId);
        Assert.Null(request.Args);
    }
}
