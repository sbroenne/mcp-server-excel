using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "DocumentInspector")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class DocumentInspectorProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Inspect_ForwardsStateChangingActionAndNativeFindings()
    {
        const string payload = """{"success":true,"inspectors":[{"index":3,"name":"Hidden &Worksheets","statusCode":1,"status":"issue-found","results":"One finding"}]}""";
        var call = await fixture.CallToolAsync("workbook", new()
        {
            ["action"] = "inspect-document",
            ["workbook_session_id"] = "session-1"
        }, RecordingToolTest.Success(payload), "workbook.inspect-document", null);
        Assert.False(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.Equal("Hidden &Worksheets", result.RootElement.GetProperty("inspectors")[0].GetProperty("name").GetString());
        Assert.Equal(1, result.RootElement.GetProperty("inspectors")[0].GetProperty("statusCode").GetInt32());
    }

    [Fact]
    public async Task Fix_ForwardsExactIndexAndLocalizedName()
    {
        var call = await fixture.CallToolAsync("workbook", new()
        {
            ["action"] = "fix-document-inspector",
            ["workbook_session_id"] = "session-1",
            ["inspector_index"] = 3,
            ["inspector_name"] = "Hidden &Worksheets"
        }, RecordingToolTest.Success("""{"success":true}"""), "workbook.fix-document-inspector",
            """{"inspectorIndex":3,"inspectorName":"Hidden &Worksheets"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task NativeFailure_IsProtocolErrorWithPartialResults()
    {
        const string payload = """{"success":false,"errorMessage":"Native inspector failed","inspectors":[{"index":1,"statusCode":2,"results":"Native error"}]}""";
        var call = await fixture.CallToolAsync("workbook", new()
        {
            ["action"] = "inspect-document",
            ["workbook_session_id"] = "session-1"
        }, RecordingToolTest.Success(payload), "workbook.inspect-document", null);
        Assert.True(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("Native error", result.RootElement.GetProperty("inspectors")[0].GetProperty("results").GetString());
    }

    [Fact]
    public async Task Fix_RequiresBothSelectionParametersBeforeDispatch()
    {
        foreach (var missing in new[] { "inspector_index", "inspector_name" })
        {
            var arguments = new Dictionary<string, object?>
            {
                ["action"] = "fix-document-inspector",
                ["workbook_session_id"] = "session-1",
                ["inspector_index"] = 3,
                ["inspector_name"] = "Hidden &Worksheets"
            };
            arguments.Remove(missing);
            var result = await fixture.CallResultWithoutDispatchAsync("workbook", arguments);
            Assert.True(result.IsError);
        }
    }

    [Fact]
    public async Task Discovery_InspectionDoesNotAdvertiseReadOnlyAndExplainsLimits()
    {
        var tools = await fixture.ListToolsAsync();
        var read = Assert.Single(tools, tool => tool.Name == "workbook_read");
        var write = Assert.Single(tools, tool => tool.Name == "workbook");
        Assert.True(read.ProtocolTool.Annotations?.ReadOnlyHint == true);
        Assert.True(write.ProtocolTool.Annotations?.ReadOnlyHint == false);
        Assert.True(write.ProtocolTool.Annotations?.DestructiveHint == true);
        var readActions = read.JsonSchema.GetProperty("properties").GetProperty("action").ToString();
        Assert.Contains("\"inspect\"", readActions, StringComparison.Ordinal);
        Assert.DoesNotContain("inspect-document", readActions, StringComparison.Ordinal);
        Assert.Contains("inspect-document", write.JsonSchema.GetProperty("properties")
            .GetProperty("action").ToString(), StringComparison.Ordinal);
        Assert.Contains("inspect-document", write.Description, StringComparison.Ordinal);
        Assert.DoesNotContain("fix-document-inspector", read.JsonSchema.GetProperty("properties")
            .GetProperty("action").ToString(), StringComparison.Ordinal);
        Assert.Contains("fix-document-inspector", write.Description, StringComparison.Ordinal);
        Assert.Contains("comments/revisions and document properties", write.Description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("unsaved", write.Description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("no tool-level undo", write.Description, StringComparison.OrdinalIgnoreCase);
        var properties = write.JsonSchema.GetProperty("properties");
        Assert.Contains(properties.GetProperty("inspector_index").GetProperty("type").EnumerateArray(),
            type => type.GetString() == "integer");
        Assert.Contains(properties.GetProperty("inspector_name").GetProperty("type").EnumerateArray(),
            type => type.GetString() == "string");
    }
}
