using System.Text.Json;
using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Analysis")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class AnalysisServiceRegistryRoutingTests
{
    public static TheoryData<AnalysisAction, string> Actions =>
        new()
        {
            { AnalysisAction.GoalSeek, "goal-seek" },
            { AnalysisAction.ListScenarios, "list-scenarios" },
            { AnalysisAction.CreateScenario, "create-scenario" },
            { AnalysisAction.UpdateScenario, "update-scenario" },
            { AnalysisAction.ShowScenario, "show-scenario" },
            { AnalysisAction.DeleteScenario, "delete-scenario" },
            { AnalysisAction.CreateScenarioSummary, "create-scenario-summary" },
            { AnalysisAction.CreateDataTable, "create-data-table" }
        };

    [Theory]
    [MemberData(nameof(Actions))]
    public void GeneratedRoutes_MapEveryAnalysisAction(AnalysisAction action, string expectedAction)
    {
        Assert.True(ServiceRegistry.Analysis.TryParseAction(expectedAction, out var parsedAction));
        Assert.Equal(action, parsedAction);
        Assert.Equal(expectedAction, ServiceRegistry.Analysis.ToActionString(action));

        var routed = ServiceRegistry.Analysis.RouteAction(
            action,
            "session-1",
            (command, sessionId, args) => JsonSerializer.Serialize(new { command, sessionId, args }),
            sheetName: "Model",
            formulaCell: "B1",
            goal: 40d,
            changingCell: "A1",
            scenarioName: "Plan",
            changingCells: "A1:A2",
            values: [10d, 20d],
            comment: "Planning inputs",
            locked: false,
            hidden: false,
            reportType: "pivot-table",
            resultCells: "B1",
            tableRange: "D1:F4",
            rowInputCell: "A1",
            columnInputCell: "A2");

        using var routedJson = JsonDocument.Parse(routed);
        Assert.Equal($"analysis.{expectedAction}", routedJson.RootElement.GetProperty("command").GetString());
        Assert.Equal("session-1", routedJson.RootElement.GetProperty("sessionId").GetString());

        var cliRoute = ServiceRegistry.Analysis.RouteCliArgs(
            expectedAction,
            sheetName: "Model",
            formulaCell: "B1",
            goal: 40d,
            changingCell: "A1",
            scenarioName: "Plan",
            changingCells: "A1:A2",
            values: [10d, 20d],
            comment: "Planning inputs",
            locked: false,
            hidden: false,
            reportType: "pivot-table",
            resultCells: "B1",
            tableRange: "D1:F4",
            rowInputCell: "A1",
            columnInputCell: "A2");

        Assert.Equal($"analysis.{expectedAction}", cliRoute.Command);
        Assert.NotNull(cliRoute.Args);
    }

    [Fact]
    public void GeneratedGoalSeekRoute_PreservesOmittedGoalForCoreValidation()
    {
        var routed = ServiceRegistry.Analysis.RouteAction(
            AnalysisAction.GoalSeek,
            "session-1",
            (_, _, args) => JsonSerializer.Serialize(args),
            sheetName: "Model",
            formulaCell: "B1",
            goal: null,
            changingCell: "A1");

        using var routedJson = JsonDocument.Parse(routed);
        Assert.Equal(JsonValueKind.Null, routedJson.RootElement.GetProperty("Goal").ValueKind);
    }
}
