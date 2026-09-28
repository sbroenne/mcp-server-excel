using Sbroenne.ExcelMcp.Core.Commands.Analysis;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Analysis")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceAnalysisTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IAnalysisCommands _analysis =
        ServiceCommandProxy.Create<IAnalysisCommands>(fixture);
    private readonly ISheetCommands _sheets =
        ServiceCommandProxy.Create<ISheetCommands>(fixture);

    [Fact]
    public void GoalSeek_AdjustsChangingCellToReachGoal()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(batch, sheetName, "A1", [[5d]]);
        _commands.SetFormulas(batch, sheetName, "B1", [["=A1*2"]]);

        var result = _analysis.GoalSeek(batch, sheetName, "B1", 40d, "A1");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.Converged);
        Assert.Equal(20d, ReadDouble(batch, sheetName, "A1"), 6);
        Assert.Equal(40d, ReadDouble(batch, sheetName, "B1"), 6);
    }

    [Fact]
    public void ScenarioLifecycle_CreateListShowUpdateDelete_ChangesRealCells()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(batch, sheetName, "A1:A2", [[1d], [2d]]);

        var createResult = _analysis.CreateScenario(
            batch, sheetName, "Best Case", "A1:A2", [10d, 20d],
            "Optimistic inputs", locked: false, hidden: false);
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var scenario = Assert.Single(_analysis.ListScenarios(batch, sheetName).Scenarios);
        Assert.Equal("Best Case", scenario.Name);
        Assert.Equal("$A$1:$A$2", scenario.ChangingCells);
        Assert.Equal([10d, 20d], scenario.Values.Select(Convert.ToDouble));

        Assert.True(_analysis.ShowScenario(batch, sheetName, "Best Case").Success);
        Assert.Equal(10d, ReadDouble(batch, sheetName, "A1"), 6);
        Assert.Equal(20d, ReadDouble(batch, sheetName, "A2"), 6);

        Assert.True(_analysis.UpdateScenario(
            batch, sheetName, "Best Case", "A1:A2", [30d, 40d]).Success);
        _analysis.ShowScenario(batch, sheetName, "Best Case");
        Assert.Equal(30d, ReadDouble(batch, sheetName, "A1"), 6);
        Assert.Equal(40d, ReadDouble(batch, sheetName, "A2"), 6);

        Assert.True(_analysis.DeleteScenario(batch, sheetName, "Best Case").Success);
        Assert.Empty(_analysis.ListScenarios(batch, sheetName).Scenarios);
    }

    [Fact]
    public void ListScenarios_ReturnsMetadata()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(batch, sheetName, "C1", [[7d]]);
        _analysis.CreateScenario(
            batch, sheetName, "Locked Plan", "C1", [11d],
            "Planning case", locked: true, hidden: true);

        var scenario = Assert.Single(_analysis.ListScenarios(batch, sheetName).Scenarios);

        Assert.Contains("Planning case", scenario.Comment, StringComparison.Ordinal);
        Assert.True(scenario.Locked);
        Assert.True(scenario.Hidden);
    }

    [Theory]
    [InlineData(ScenarioSummaryType.Summary)]
    [InlineData(ScenarioSummaryType.PivotTable)]
    public void CreateScenarioSummary_AddsReportWorksheet(ScenarioSummaryType reportType)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(batch, sheetName, "A1", [[2d]]);
        _commands.SetFormulas(batch, sheetName, "B1", [["=A1*3"]]);
        _analysis.CreateScenario(batch, sheetName, "Base", "A1", [2d]);
        _analysis.CreateScenario(batch, sheetName, "Growth", "A1", [5d]);
        var sheetCountBefore = _sheets.List(batch).Worksheets.Count;

        var result = _analysis.CreateScenarioSummary(
            batch, sheetName, reportType, "B1");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(result.ReportSheetName));
        _fixture.RegisterSheetForCleanup(result.ReportSheetName);
        Assert.Equal(sheetCountBefore + 1, _sheets.List(batch).Worksheets.Count);
        Assert.Equal(
            reportType == ScenarioSummaryType.Summary ? "summary" : "pivot-table",
            result.ReportType);
    }

    [Fact]
    public void CreateDataTable_OneVariable_PopulatesResults()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(
            batch, sheetName, "A1:D4",
            [[null, null, null, 0d], [1d, null, null, null],
             [2d, null, null, null], [3d, null, null, null]]);
        _commands.SetFormulas(batch, sheetName, "B1", [["=D1*2"]]);

        var result = _analysis.CreateDataTable(
            batch, sheetName, "A1:B4", columnInputCell: "D1");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2d, ReadDouble(batch, sheetName, "B2"), 6);
        Assert.Equal(4d, ReadDouble(batch, sheetName, "B3"), 6);
        Assert.Equal(6d, ReadDouble(batch, sheetName, "B4"), 6);
    }

    [Fact]
    public void CreateDataTable_TwoVariable_PopulatesMatrix()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(
            batch, sheetName, "A1:C13",
            [[null, 2d, 3d], [4d, null, null], [5d, null, null],
             [null, null, null], [null, null, null], [null, null, null],
             [null, null, null], [null, null, null], [null, null, null],
             [null, null, null], [null, null, null], [0d, null, null],
             [0d, null, null]]);
        _commands.SetFormulas(batch, sheetName, "A1", [["=A12*100+A13"]]);

        var result = _analysis.CreateDataTable(
            batch, sheetName, "A1:C3", "A12", "A13");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(204d, ReadDouble(batch, sheetName, "B2"), 6);
        Assert.Equal(305d, ReadDouble(batch, sheetName, "C3"), 6);
    }

    private double ReadDouble(
        Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch,
        string sheetName,
        string address) =>
        Convert.ToDouble(
            _commands.GetValues(batch, sheetName, address).Values[0][0],
            System.Globalization.CultureInfo.InvariantCulture);
}
