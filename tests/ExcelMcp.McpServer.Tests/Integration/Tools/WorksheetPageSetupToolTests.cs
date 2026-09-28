using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "false")]
public sealed class WorksheetPageSetupToolTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task WorksheetStyle_SetPageSetup_RoundsTripThroughMcp()
    {
        const string sessionId = "recording-session";
        var setCall = await _fixture.CallToolAsync(
            "worksheet_style",
            new Dictionary<string, object?>
            {
                ["action"] = "set-page-setup",
                ["session_id"] = sessionId,
                ["sheet_name"] = "PageSetupSheet",
                ["orientation"] = "landscape",
                ["fit_to_pages_wide"] = 1,
                ["fit_to_pages_tall"] = 2,
                ["center_horizontally"] = false,
                ["center_vertically"] = true
            },
            Success("""{"success":true}"""),
            "sheet.set-page-setup",
            """{"sheetName":"PageSetupSheet","orientation":"landscape","fitToPagesWide":1,"fitToPagesTall":2,"centerHorizontally":false,"centerVertically":true}""");

        Assert.Equal("sheet.set-page-setup", setCall.Request.Command);
        Assert.Equal(sessionId, setCall.Request.SessionId);
        using (var args = ParseArgs(setCall.Request))
        {
            var root = args.RootElement;
            Assert.Equal("PageSetupSheet", root.GetProperty("sheetName").GetString());
            Assert.Equal("landscape", root.GetProperty("orientation").GetString());
            Assert.Equal(1, root.GetProperty("fitToPagesWide").GetInt32());
            Assert.Equal(2, root.GetProperty("fitToPagesTall").GetInt32());
            Assert.False(root.GetProperty("centerHorizontally").GetBoolean());
            Assert.True(root.GetProperty("centerVertically").GetBoolean());
        }

        var getCall = await _fixture.CallToolAsync(
            "worksheet_style",
            new Dictionary<string, object?>
            {
                ["action"] = "get-page-setup",
                ["session_id"] = sessionId,
                ["sheet_name"] = "PageSetupSheet"
            },
            Success(
                """{"success":true,"orientation":"landscape","fitToPagesWide":1,"fitToPagesTall":2,"centerHorizontally":false,"centerVertically":true}"""),
            "sheet.get-page-setup",
            """{"sheetName":"PageSetupSheet"}""");

        Assert.Equal("sheet.get-page-setup", getCall.Request.Command);
        using var result = JsonDocument.Parse(getCall.JsonResult);
        var resultRoot = result.RootElement;
        Assert.Equal("landscape", resultRoot.GetProperty("orientation").GetString());
        Assert.Equal(1, resultRoot.GetProperty("fitToPagesWide").GetInt32());
        Assert.Equal(2, resultRoot.GetProperty("fitToPagesTall").GetInt32());
        Assert.False(resultRoot.GetProperty("centerHorizontally").GetBoolean());
        Assert.True(resultRoot.GetProperty("centerVertically").GetBoolean());
    }

    [Fact]
    public async Task WorksheetStyle_GetPageSetup_AutomaticScalingReturnsNullFitValues()
    {
        const string sessionId = "recording-session";
        var call = await _fixture.CallToolAsync(
            "worksheet_style",
            new Dictionary<string, object?>
            {
                ["action"] = "get-page-setup",
                ["session_id"] = sessionId,
                ["sheet_name"] = "AutomaticScale"
            },
            Success(
                """{"success":true,"fitToPagesWide":null,"fitToPagesTall":null}"""),
            "sheet.get-page-setup",
            """{"sheetName":"AutomaticScale"}""");

        Assert.Equal("sheet.get-page-setup", call.Request.Command);
        Assert.Equal(sessionId, call.Request.SessionId);
        using (var args = ParseArgs(call.Request))
        {
            Assert.Equal(
                "AutomaticScale",
                args.RootElement.GetProperty("sheetName").GetString());
        }

        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.Equal(
            JsonValueKind.Null,
            result.RootElement.GetProperty("fitToPagesWide").ValueKind);
        Assert.Equal(
            JsonValueKind.Null,
            result.RootElement.GetProperty("fitToPagesTall").ValueKind);
    }

    private static ServiceResponse Success(string result) =>
        new()
        {
            Success = true,
            Result = result
        };

    private static JsonDocument ParseArgs(ServiceRequest request)
    {
        Assert.NotNull(request.Args);
        return JsonDocument.Parse(request.Args);
    }
}
