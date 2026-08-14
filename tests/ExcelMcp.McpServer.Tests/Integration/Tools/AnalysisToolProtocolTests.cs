using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// End-to-end MCP protocol coverage for Excel what-if analysis.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Analysis")]
[Trait("RequiresExcel", "true")]
public sealed class AnalysisToolProtocolTests : McpIntegrationTestBase
{
    private readonly string _tempDir;

    public AnalysisToolProtocolTests(ITestOutputHelper output)
        : base(output, "AnalysisToolProtocolClient")
    {
        _tempDir = CreateTempDirectory("AnalysisToolProtocolTests");
    }

    [Fact]
    public async Task GoalSeek_AdjustsChangingCellThroughMcp()
    {
        var sessionId = await CreateWorkbookSessionAsync(Path.Join(_tempDir, $"GoalSeek_{Guid.NewGuid():N}.xlsx"));
        await SetValuesAsync(sessionId, "A1", [[5d]]);
        await SetFormulasAsync(sessionId, "B1", [["=A1*2"]]);

        var result = await CallToolAsync("analysis", new Dictionary<string, object?>
        {
            ["action"] = "goal-seek",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["formula_cell"] = "B1",
            ["goal"] = 40d,
            ["changing_cell"] = "A1"
        });

        AssertSuccess(result, "analysis.goal-seek");
        using var resultJson = JsonDocument.Parse(result);
        Assert.True(resultJson.RootElement.GetProperty("converged").GetBoolean());
        Assert.Equal(20d, await ReadSingleValueAsync(sessionId, "A1"), 6);
        await CloseSessionAsync(sessionId, save: false);
    }

    [Fact]
    public async Task GoalSeek_MissingGoal_ReturnsTransparentFailureThroughMcp()
    {
        var sessionId = await CreateWorkbookSessionAsync(Path.Join(_tempDir, $"GoalSeekMissingGoal_{Guid.NewGuid():N}.xlsx"));
        await SetValuesAsync(sessionId, "A1", [[5d]]);
        await SetFormulasAsync(sessionId, "B1", [["=A1*2"]]);

        var result = await CallToolAsync("analysis", new Dictionary<string, object?>
        {
            ["action"] = "goal-seek",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["formula_cell"] = "B1",
            ["changing_cell"] = "A1"
        });

        using var resultJson = ParseJsonResult(result, "analysis.goal-seek missing goal");
        AssertFailureEnvelope(
            resultJson.RootElement,
            "analysis.goal-seek missing goal",
            nameof(ArgumentNullException),
            expectedErrorCategory: "InvalidInput");
        Assert.Contains(
            "goal",
            resultJson.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
        await CloseSessionAsync(sessionId, save: false);
    }

    [Fact]
    public async Task ScenarioLifecycleAndSummary_WorkThroughMcp()
    {
        var sessionId = await CreateWorkbookSessionAsync(Path.Join(_tempDir, $"Scenarios_{Guid.NewGuid():N}.xlsx"));
        await SetValuesAsync(sessionId, "A1:A2", [[1d], [2d]]);
        await SetFormulasAsync(sessionId, "B1", [["=SUM(A1:A2)"]]);

        var create = await CallToolAsync("analysis", new Dictionary<string, object?>
        {
            ["action"] = "create-scenario",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["scenario_name"] = "Plan",
            ["changing_cells"] = "A1:A2",
            ["values"] = new List<object?> { 10d, 20d },
            ["comment"] = "MCP scenario",
            ["locked"] = false,
            ["hidden"] = false
        });
        AssertSuccess(create, "analysis.create-scenario");

        var list = await CallToolAsync("analysis", new Dictionary<string, object?>
        {
            ["action"] = "list-scenarios",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1"
        });
        using (var listJson = JsonDocument.Parse(list))
        {
            var scenario = Assert.Single(listJson.RootElement.GetProperty("scenarios").EnumerateArray());
            Assert.Equal("Plan", scenario.GetProperty("name").GetString());
            Assert.Equal(2, scenario.GetProperty("values").GetArrayLength());
        }

        var update = await CallToolAsync("analysis", new Dictionary<string, object?>
        {
            ["action"] = "update-scenario",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["scenario_name"] = "Plan",
            ["changing_cells"] = "A1:A2",
            ["values"] = new List<object?> { 30d, 40d }
        });
        AssertSuccess(update, "analysis.update-scenario");

        var show = await CallToolAsync("analysis", new Dictionary<string, object?>
        {
            ["action"] = "show-scenario",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["scenario_name"] = "Plan"
        });
        AssertSuccess(show, "analysis.show-scenario");
        Assert.Equal(30d, await ReadSingleValueAsync(sessionId, "A1"), 6);

        var summary = await CallToolAsync("analysis", new Dictionary<string, object?>
        {
            ["action"] = "create-scenario-summary",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["report_type"] = "pivot-table",
            ["result_cells"] = "B1"
        });
        AssertSuccess(summary, "analysis.create-scenario-summary");
        using (var summaryJson = JsonDocument.Parse(summary))
        {
            Assert.False(string.IsNullOrWhiteSpace(summaryJson.RootElement.GetProperty("reportSheetName").GetString()));
            Assert.Equal("pivot-table", summaryJson.RootElement.GetProperty("reportType").GetString());
        }

        var standardSummary = await CallToolAsync("analysis", new Dictionary<string, object?>
        {
            ["action"] = "create-scenario-summary",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["report_type"] = "summary",
            ["result_cells"] = "B1"
        });
        AssertSuccess(standardSummary, "analysis.create-scenario-summary standard");
        using (var summaryJson = JsonDocument.Parse(standardSummary))
        {
            Assert.False(string.IsNullOrWhiteSpace(summaryJson.RootElement.GetProperty("reportSheetName").GetString()));
            Assert.Equal("summary", summaryJson.RootElement.GetProperty("reportType").GetString());
        }

        var delete = await CallToolAsync("analysis", new Dictionary<string, object?>
        {
            ["action"] = "delete-scenario",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["scenario_name"] = "Plan"
        });
        AssertSuccess(delete, "analysis.delete-scenario");
        await CloseSessionAsync(sessionId, save: false);
    }

    [Fact]
    public async Task CreateDataTable_PopulatesResultsThroughMcp()
    {
        var sessionId = await CreateWorkbookSessionAsync(Path.Join(_tempDir, $"DataTable_{Guid.NewGuid():N}.xlsx"));
        await SetValuesAsync(sessionId, "A2:A4", [[1d], [2d], [3d]]);
        await SetValuesAsync(sessionId, "D1", [[0d]]);
        await SetFormulasAsync(sessionId, "B1", [["=D1*2"]]);

        var result = await CallToolAsync("analysis", new Dictionary<string, object?>
        {
            ["action"] = "create-data-table",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["table_range"] = "A1:B4",
            ["column_input_cell"] = "D1"
        });

        AssertSuccess(result, "analysis.create-data-table");
        Assert.Equal(6d, await ReadSingleValueAsync(sessionId, "B4"), 6);
        await CloseSessionAsync(sessionId, save: false);
    }

    [Fact]
    public async Task CreateDataTable_TwoVariablePreservesInputArgumentOrderThroughMcp()
    {
        var sessionId = await CreateWorkbookSessionAsync(Path.Join(_tempDir, $"TwoVariableDataTable_{Guid.NewGuid():N}.xlsx"));
        await SetValuesAsync(sessionId, "A12:A13", [[0d], [0d]]);
        await SetFormulasAsync(sessionId, "A1", [["=A12*100+A13"]]);
        await SetValuesAsync(sessionId, "B1:C1", [[2d, 3d]]);
        await SetValuesAsync(sessionId, "A2:A3", [[4d], [5d]]);

        var result = await CallToolAsync("analysis", new Dictionary<string, object?>
        {
            ["action"] = "create-data-table",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["table_range"] = "A1:C3",
            ["row_input_cell"] = "A12",
            ["column_input_cell"] = "A13"
        });

        AssertSuccess(result, "analysis.create-data-table two-variable");
        Assert.Equal(204d, await ReadSingleValueAsync(sessionId, "B2"), 6);
        Assert.Equal(305d, await ReadSingleValueAsync(sessionId, "C3"), 6);
        await CloseSessionAsync(sessionId, save: false);
    }

    private async Task SetValuesAsync(string sessionId, string rangeAddress, List<List<object?>> values)
    {
        var result = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = rangeAddress,
            ["values"] = values
        });
        AssertSetupSuccess(result, $"range.set-values ({rangeAddress})");
    }

    private async Task SetFormulasAsync(string sessionId, string rangeAddress, List<List<string>> formulas)
    {
        var result = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-formulas",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = rangeAddress,
            ["formulas"] = formulas
        });
        AssertSetupSuccess(result, $"range.set-formulas ({rangeAddress})");
    }

    private async Task<double> ReadSingleValueAsync(string sessionId, string rangeAddress)
    {
        var result = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = rangeAddress
        });
        AssertSetupSuccess(result, $"range.get-values ({rangeAddress})");
        using var json = JsonDocument.Parse(result);
        return json.RootElement.GetProperty("values")[0][0].GetDouble();
    }
}
