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
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[5d]]).Success);
        Assert.True(_commands.SetFormulas(batch, sheetName, "B1", [["=A1*2"]]).Success);
        Assert.Equal(5d, ReadDouble(batch, sheetName, "A1"), 6);
        Assert.Equal(10d, ReadDouble(batch, sheetName, "B1"), 6);

        var result = _analysis.GoalSeek(batch, sheetName, "B1", 40d, "A1");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.Converged);
        Assert.Equal(20d, ReadDouble(batch, sheetName, "A1"), 6);
        Assert.Equal(40d, ReadDouble(batch, sheetName, "B1"), 6);
        var formula = _commands.GetFormulas(batch, sheetName, "B1");
        Assert.True(formula.Success, formula.ErrorMessage);
        Assert.Equal("=A1*2", Assert.Single(Assert.Single(formula.Formulas)));
    }

    [Fact]
    public void ScenarioLifecycle_CreateListShowUpdateDelete_ChangesRealCells()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:A2", [[1d], [2d]]).Success);

        var createResult = _analysis.CreateScenario(
            batch, sheetName, "Best Case", "A1:A2", [10d, 20d],
            "Optimistic inputs", locked: false, hidden: false);
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var listed = _analysis.ListScenarios(batch, sheetName);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal(sheetName, listed.SheetName);
        var scenario = Assert.Single(listed.Scenarios);
        Assert.Equal("Best Case", scenario.Name);
        Assert.Equal("$A$1:$A$2", scenario.ChangingCells);
        Assert.Equal([10d, 20d], scenario.Values.Select(Convert.ToDouble));

        Assert.True(_analysis.ShowScenario(batch, sheetName, "Best Case").Success);
        Assert.Equal(10d, ReadDouble(batch, sheetName, "A1"), 6);
        Assert.Equal(20d, ReadDouble(batch, sheetName, "A2"), 6);

        Assert.True(_analysis.UpdateScenario(
            batch, sheetName, "Best Case", "A1:A2", [30d, 40d]).Success);
        Assert.True(_analysis.ShowScenario(batch, sheetName, "Best Case").Success);
        Assert.Equal(30d, ReadDouble(batch, sheetName, "A1"), 6);
        Assert.Equal(40d, ReadDouble(batch, sheetName, "A2"), 6);

        Assert.True(_analysis.DeleteScenario(batch, sheetName, "Best Case").Success);
        var afterDelete = _analysis.ListScenarios(batch, sheetName);
        Assert.True(afterDelete.Success, afterDelete.ErrorMessage);
        Assert.Empty(afterDelete.Scenarios);
        Assert.Equal(30d, ReadDouble(batch, sheetName, "A1"), 6);
        Assert.Equal(40d, ReadDouble(batch, sheetName, "A2"), 6);
    }

    [Fact]
    public void ListScenarios_ReturnsMetadata()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "C1", [[7d]]).Success);
        var created = _analysis.CreateScenario(
            batch, sheetName, "Locked Plan", "C1", [11d],
            "Planning case", locked: true, hidden: true);
        Assert.True(created.Success, created.ErrorMessage);

        var listed = _analysis.ListScenarios(batch, sheetName);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal(sheetName, listed.SheetName);
        var scenario = Assert.Single(listed.Scenarios);

        Assert.Equal("Locked Plan", scenario.Name);
        Assert.Equal("$C$1", scenario.ChangingCells);
        Assert.Equal(11d, Convert.ToDouble(Assert.Single(scenario.Values), System.Globalization.CultureInfo.InvariantCulture));
        Assert.Contains("Planning case", scenario.Comment, StringComparison.Ordinal);
        Assert.True(scenario.Locked);
        Assert.True(scenario.Hidden);
        Assert.Equal(7d, ReadDouble(batch, sheetName, "C1"), 6);
    }

    [Theory]
    [InlineData(ScenarioSummaryType.Summary)]
    [InlineData(ScenarioSummaryType.PivotTable)]
    public void CreateScenarioSummary_AddsReportWorksheet(ScenarioSummaryType reportType)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[2d]]).Success);
        Assert.True(_commands.SetFormulas(batch, sheetName, "B1", [["=A1*3"]]).Success);
        Assert.True(_analysis.CreateScenario(batch, sheetName, "Base", "A1", [2d]).Success);
        Assert.True(_analysis.CreateScenario(batch, sheetName, "Growth", "A1", [5d]).Success);
        var before = _sheets.List(batch);
        Assert.True(before.Success, before.ErrorMessage);
        var sheetCountBefore = before.Worksheets.Count;

        var result = _analysis.CreateScenarioSummary(
            batch, sheetName, reportType, "B1");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(result.ReportSheetName));
        _fixture.RegisterSheetForCleanup(result.ReportSheetName);
        var after = _sheets.List(batch);
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(sheetCountBefore + 1, after.Worksheets.Count);
        Assert.Equal(
            reportType == ScenarioSummaryType.Summary ? "summary" : "pivot-table",
            result.ReportType);
        var report = _commands.GetUsedRange(batch, result.ReportSheetName);
        Assert.True(report.Success, report.ErrorMessage);
        var cells = report.Values.SelectMany(row => row).ToList();
        Assert.Contains("Base", cells);
        Assert.Contains("Growth", cells);
        var amounts = cells.Where(value => value is int or long or double or decimal)
            .Select(value => Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture));
        Assert.Contains(6d, amounts);
        Assert.Contains(15d, amounts);
        var baseCell = FindReportCell("Base");
        var growthCell = FindReportCell("Growth");
        if (baseCell.Row == growthCell.Row)
        {
            Assert.NotEqual(baseCell.Column, growthCell.Column);
            Assert.Single(report.Values, row =>
                IsAmount(row[baseCell.Column], 6d) && IsAmount(row[growthCell.Column], 15d));
        }
        else
        {
            Assert.Equal(baseCell.Column, growthCell.Column);
            Assert.Single(Enumerable.Range(0, report.ColumnCount), column =>
                IsAmount(report.Values[baseCell.Row][column], 6d) &&
                IsAmount(report.Values[growthCell.Row][column], 15d));
        }

        (int Row, int Column) FindReportCell(string label) =>
            Assert.Single(report.Values.SelectMany((row, rowIndex) =>
                row.Select((value, columnIndex) => (value, rowIndex, columnIndex)))
                .Where(cell => Equals(cell.value, label))
                .Select(cell => (cell.rowIndex, cell.columnIndex)));

        static bool IsAmount(object? value, double expected) =>
            value is int or long or double or decimal &&
            Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture) == expected;
    }

    [Fact]
    public void CreateDataTable_OneVariable_PopulatesResults()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(
            batch, sheetName, "A1:D4",
            [[null, null, null, 0d], [1d, null, null, null],
             [2d, null, null, null], [3d, null, null, null]]).Success);
        Assert.True(_commands.SetFormulas(batch, sheetName, "B1", [["=D1*2"]]).Success);

        var result = _analysis.CreateDataTable(
            batch, sheetName, "A1:B4", columnInputCell: "D1");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2d, ReadDouble(batch, sheetName, "B2"), 6);
        Assert.Equal(4d, ReadDouble(batch, sheetName, "B3"), 6);
        Assert.Equal(6d, ReadDouble(batch, sheetName, "B4"), 6);
        Assert.Equal(0d, ReadDouble(batch, sheetName, "D1"), 6);
        Assert.Equal(1d, ReadDouble(batch, sheetName, "A2"), 6);
        Assert.Equal(2d, ReadDouble(batch, sheetName, "A3"), 6);
        Assert.Equal(3d, ReadDouble(batch, sheetName, "A4"), 6);
    }

    [Fact]
    public void CreateDataTable_TwoVariable_PopulatesMatrix()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(
            batch, sheetName, "A1:C13",
            [[null, 2d, 3d], [4d, null, null], [5d, null, null],
             [null, null, null], [null, null, null], [null, null, null],
             [null, null, null], [null, null, null], [null, null, null],
             [null, null, null], [null, null, null], [0d, null, null],
             [0d, null, null]]).Success);
        Assert.True(_commands.SetFormulas(batch, sheetName, "A1", [["=A12*100+A13"]]).Success);

        var result = _analysis.CreateDataTable(
            batch, sheetName, "A1:C3", "A12", "A13");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(204d, ReadDouble(batch, sheetName, "B2"), 6);
        Assert.Equal(304d, ReadDouble(batch, sheetName, "C2"), 6);
        Assert.Equal(205d, ReadDouble(batch, sheetName, "B3"), 6);
        Assert.Equal(305d, ReadDouble(batch, sheetName, "C3"), 6);
        Assert.Equal(0d, ReadDouble(batch, sheetName, "A12"), 6);
        Assert.Equal(0d, ReadDouble(batch, sheetName, "A13"), 6);
        Assert.Equal(2d, ReadDouble(batch, sheetName, "B1"), 6);
        Assert.Equal(3d, ReadDouble(batch, sheetName, "C1"), 6);
        Assert.Equal(4d, ReadDouble(batch, sheetName, "A2"), 6);
        Assert.Equal(5d, ReadDouble(batch, sheetName, "A3"), 6);
    }

    [Fact]
    public void UpdateScenario_MismatchedValues_PreservesScenarioAndCells()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:A2", [[1d], [2d]]).Success);
        var created = _analysis.CreateScenario(batch, sheetName, "Retained", "A1:A2", [10d, 20d]);
        Assert.True(created.Success, created.ErrorMessage);
        var before = _analysis.ListScenarios(batch, sheetName);
        Assert.True(before.Success, before.ErrorMessage);
        var original = Assert.Single(before.Scenarios);
        Assert.Equal("Retained", original.Name);
        Assert.Equal("$A$1:$A$2", original.ChangingCells);
        Assert.Equal([10d, 20d], original.Values.Select(Convert.ToDouble));

        var exception = Assert.Throws<ArgumentException>(() =>
            _analysis.UpdateScenario(batch, sheetName, "Retained", "A1:A2", [99d]));

        Assert.Contains("must match changing cells count", exception.Message, StringComparison.Ordinal);
        var listed = _analysis.ListScenarios(batch, sheetName);
        Assert.True(listed.Success, listed.ErrorMessage);
        var retained = Assert.Single(listed.Scenarios);
        Assert.Equal(original.Name, retained.Name);
        Assert.Equal(original.ChangingCells, retained.ChangingCells);
        Assert.Equal(original.Comment, retained.Comment);
        Assert.Equal(original.Locked, retained.Locked);
        Assert.Equal(original.Hidden, retained.Hidden);
        Assert.Equal([10d, 20d], retained.Values.Select(Convert.ToDouble));
        Assert.Equal(1d, ReadDouble(batch, sheetName, "A1"), 6);
        Assert.Equal(2d, ReadDouble(batch, sheetName, "A2"), 6);
    }

    [Fact]
    public void CreateScenario_MismatchedValues_PreservesExistingScenarioAndCells()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:A2", [[1d], [2d]]).Success);
        Assert.True(_analysis.CreateScenario(batch, sheetName, "Retained", "A1:A2", [10d, 20d]).Success);
        var before = _analysis.ListScenarios(batch, sheetName);
        Assert.True(before.Success, before.ErrorMessage);
        var original = Assert.Single(before.Scenarios);

        var exception = Assert.Throws<ArgumentException>(() =>
            _analysis.CreateScenario(batch, sheetName, "Rejected", "A1:A2", [99d]));

        Assert.Contains("must match changing cells count", exception.Message, StringComparison.Ordinal);
        var after = _analysis.ListScenarios(batch, sheetName);
        Assert.True(after.Success, after.ErrorMessage);
        var retained = Assert.Single(after.Scenarios);
        Assert.Equal("Retained", retained.Name);
        Assert.Equal(original.ChangingCells, retained.ChangingCells);
        Assert.Equal(original.Comment, retained.Comment);
        Assert.Equal(original.Locked, retained.Locked);
        Assert.Equal(original.Hidden, retained.Hidden);
        Assert.Equal([10d, 20d], retained.Values.Select(Convert.ToDouble));
        Assert.Equal(1d, ReadDouble(batch, sheetName, "A1"), 6);
        Assert.Equal(2d, ReadDouble(batch, sheetName, "A2"), 6);
    }

    private double ReadDouble(
        Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch,
        string sheetName,
        string address)
    {
        var result = _commands.GetValues(batch, sheetName, address);
        Assert.True(result.Success, result.ErrorMessage);
        return Convert.ToDouble(Assert.Single(Assert.Single(result.Values)), System.Globalization.CultureInfo.InvariantCulture);
    }
}
