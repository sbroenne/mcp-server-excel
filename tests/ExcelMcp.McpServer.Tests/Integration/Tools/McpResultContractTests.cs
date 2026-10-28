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
    private static readonly string[] ValueFieldPaths = ["values"];
    private static readonly string[] QueryPreviewPaths = ["queries.name", "queries.formulaPreview"];

    [Fact]
    public async Task ReadSelection_PreservesGridDiagnosticsAndMatchingContent()
    {
        var call = await fixture.CallToolAsync("range_read", new()
        {
            ["action"] = "get-values",
            ["workbook_session_id"] = "recording-session",
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:D1",
            ["fields"] = ValueFieldPaths
        }, new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"filePath":"C:\\fixture.xlsx","sheetName":"Data","rangeAddress":"$A$1:$D$1","rowCount":1,"columnCount":4,"values":[[false,0,null,"#REF!"]],"cellErrors":[{"cellAddress":"D1","errorName":"#REF!"}]}"""
        }, "range.get-values", """{"sheetName":"Data","rangeAddress":"A1:D1","fields":["values"]}""");
        Assert.False(call.Result.IsError == true);
        using var output = JsonDocument.Parse(call.JsonResult);
        var root = output.RootElement;
        Assert.True(JsonElement.DeepEquals(root, call.Result.StructuredContent!.Value));
        Assert.Equal("""[[false,0,null,"#REF!"]]""", root.GetProperty("values").GetRawText());
        Assert.Equal(4, root.GetProperty("columnCount").GetInt32());
        Assert.Single(root.GetProperty("cellErrors").EnumerateArray());
        Assert.False(root.TryGetProperty("filePath", out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task QueryList_PreviewsRequireExplicitSelection(bool selectPreview)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["workbook_session_id"] = "recording-session"
        };
        if (selectPreview)
            arguments["fields"] = QueryPreviewPaths;
        var call = await fixture.CallToolAsync("powerquery_read", arguments, new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"queries":[{"name":"Source","formulaPreview":"let x = 1 in x","loadMode":"connection-only"}]}"""
        }, "powerquery.list",
            selectPreview ? """{"fields":["queries.name","queries.formulaPreview"]}""" : null);
        using var output = JsonDocument.Parse(call.JsonResult);
        var query = Assert.Single(output.RootElement.GetProperty("queries").EnumerateArray());
        Assert.Equal("Source", query.GetProperty("name").GetString());
        Assert.Equal(selectPreview, query.TryGetProperty("formulaPreview", out _));
        Assert.False(call.Result.IsError == true);
    }

    [Theory]
    [InlineData("queries.unknown")]
    [InlineData("Queries.name")]
    public async Task UnknownSelection_IsRejectedBeforeDispatch(string field)
    {
        var result = await fixture.CallResultWithoutDispatchAsync("powerquery_read", new()
        {
            ["action"] = "list",
            ["workbook_session_id"] = "recording-session",
            ["fields"] = new[] { field }
        });
        Assert.True(result.IsError);
        var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
        Assert.Contains("queries.name", text);
        using var output = JsonDocument.Parse(text);
        Assert.True(JsonElement.DeepEquals(output.RootElement, result.StructuredContent!.Value));
    }

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
    [InlineData("file", """{"action":"open","file_path":"C:\\missing.xlsx","show":"yes"}""", "show")]
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
