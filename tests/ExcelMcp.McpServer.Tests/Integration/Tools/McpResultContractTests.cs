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
public sealed class McpResultContractTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Failure_SetsProtocolErrorAndMatchingStructuredContent(bool dispatched)
    {
        var call = await fixture.CallToolAsync(
            "range_read",
            new Dictionary<string, object?>
            {
                ["action"] = "get-values",
                ["workbook_session_id"] = "recording-session",
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1"
            },
            new ServiceResponse
            {
                Success = dispatched,
                ErrorMessage = dispatched ? null : "Synthetic failure.",
                Result = dispatched ? """{"success":false,"errorMessage":"Synthetic failure."}""" : null
            },
            "range.get-values",
            """{"sheetName":"Sheet1","rangeAddress":"A1"}""");

        Assert.True(call.Result.IsError);
        Assert.NotNull(call.Result.StructuredContent);
        using var text = JsonDocument.Parse(call.JsonResult);
        Assert.True(JsonElement.DeepEquals(text.RootElement, call.Result.StructuredContent.Value));
    }

    [Theory]
    [InlineData("range_read", """{"action":"get-values","workbook_session_id":"s","sheet_name":"Sheet1","range_address":"A1","rang_address":"A2"}""", "rang_address")]
    [InlineData("file", """{"action":"close","workbook_session_id":"s","save_changes":true}""", "save_changes")]
    [InlineData("worksheet", """{"action":"create","workbook_session_id":"s","sheet_name":"New","before_sheet":"Sheet1"}""", "before_sheet")]
    [InlineData("file_read", """{"action":"list","save":false}""", "save")]
    [InlineData("file", """{"action":"open","path":"C:\\missing.xlsx","show":"yes"}""", "show")]
    [InlineData("file", """{}""", "action")]
    [InlineData("file", """{"action":"not-an-action"}""", "action")]
    public async Task InvalidInput_IsRejectedBeforeDispatchWithUsefulError(
        string tool, string argumentsJson, string parameter)
    {
        var arguments = JsonSerializer.Deserialize<Dictionary<string, object?>>(argumentsJson)!;
        var result = await fixture.CallResultWithoutDispatchAsync(tool, arguments);

        Assert.True(result.IsError);
        var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
        Assert.Contains(parameter, text, StringComparison.Ordinal);
        Assert.NotNull(result.StructuredContent);
    }

    [Theory]
    [InlineData("workbook_read", "get-info")]
    [InlineData("worksheet_read", "list")]
    [InlineData("screenshot", "capture")]
    [InlineData("file", "close")]
    public async Task OmittedWorkbookSessionId_IsRejectedBeforeDispatch(string tool, string action)
    {
        var result = await fixture.CallResultWithoutDispatchAsync(tool, new()
        {
            ["action"] = action
        });

        Assert.True(result.IsError);
        var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
        using var document = JsonDocument.Parse(text);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("InvalidInput", document.RootElement.GetProperty("errorCategory").GetString());
        Assert.Contains("workbook_session_id", document.RootElement.GetProperty("errorMessage").GetString());
        Assert.NotNull(result.StructuredContent);
        Assert.True(JsonElement.DeepEquals(document.RootElement, result.StructuredContent.Value));
    }

    [Theory]
    [InlineData("workbook_read", "get-info", "")]
    [InlineData("workbook_read", "get-info", " \t\r\n")]
    [InlineData("worksheet_read", "list", "")]
    [InlineData("worksheet_read", "list", " \t\r\n")]
    [InlineData("screenshot", "capture", "")]
    [InlineData("screenshot", "capture", " \t\r\n")]
    public async Task BlankWorkbookSessionId_IsRejectedBeforeDispatch(string tool, string action, string sessionId)
    {
        var result = await fixture.CallResultWithoutDispatchAsync(tool, new()
        {
            ["action"] = action,
            ["workbook_session_id"] = sessionId
        });

        Assert.True(result.IsError);
        var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
        using var document = JsonDocument.Parse(text);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("InvalidInput", document.RootElement.GetProperty("errorCategory").GetString());
        Assert.Contains("workbook_session_id", document.RootElement.GetProperty("errorMessage").GetString());
        Assert.NotNull(result.StructuredContent);
        Assert.True(JsonElement.DeepEquals(document.RootElement, result.StructuredContent.Value));
    }
}
