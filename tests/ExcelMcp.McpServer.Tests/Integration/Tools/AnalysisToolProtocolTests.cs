using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Analysis")]
[Trait("RequiresExcel", "false")]
public sealed class AnalysisToolProtocolTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task GoalSeek_MissingGoal_ReturnsTransparentFailureThroughMcp()
    {
        const string sessionId = "recording-session";
        var resultText = await _fixture.CallToolWithoutDispatchAsync(
            "analysis",
            new Dictionary<string, object?>
            {
                ["action"] = "goal-seek",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Analysis",
                ["formula_cell"] = "B1",
                ["changing_cell"] = "A1"
            });

        using var result = JsonDocument.Parse(resultText);
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(
            "ArgumentException",
            result.RootElement.GetProperty("exceptionType").GetString());
        Assert.Equal(
            "InvalidInput",
            result.RootElement.GetProperty("errorCategory").GetString());
    }

    [Fact]
    public async Task CreateDataTable_TwoVariablePreservesInputArgumentOrderThroughMcp()
    {
        const string sessionId = "recording-session";
        var call = await _fixture.CallToolAsync(
            "analysis",
            new Dictionary<string, object?>
            {
                ["action"] = "create-data-table",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Analysis",
                ["table_range"] = "A1:C3",
                ["row_input_cell"] = "A12",
                ["column_input_cell"] = "A13"
            },
            RecordingToolTest.Success("""{"success":true}"""),
            "analysis.create-data-table",
            """{"sheetName":"Analysis","tableRange":"A1:C3","rowInputCell":"A12","columnInputCell":"A13"}""");

        using var args = RecordingToolTest.ParseArgs(
            call.Request,
            "analysis.create-data-table",
            sessionId);
        var root = args.RootElement;
        Assert.Equal("A1:C3", root.GetProperty("tableRange").GetString());
        Assert.Equal("A12", root.GetProperty("rowInputCell").GetString());
        Assert.Equal("A13", root.GetProperty("columnInputCell").GetString());
    }
}
