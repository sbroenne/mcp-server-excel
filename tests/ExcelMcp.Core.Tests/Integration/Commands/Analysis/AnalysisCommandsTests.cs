using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Analysis;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Analysis;

/// <summary>
/// Integration coverage for Excel what-if analysis through the real COM object model.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Analysis")]
[Trait("RequiresExcel", "true")]
public sealed class AnalysisCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly AnalysisCommands _commands = new();
    private readonly TempDirectoryFixture _fixture;

    public AnalysisCommandsTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void GoalSeek_AdjustsChangingCellToReachGoal()
    {
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);
        ConfigureSheet(batch, sheet =>
        {
            SetValue(sheet, "A1", 5d);
            SetFormula(sheet, "B1", "=A1*2");
        });

        var result = _commands.GoalSeek(batch, "Sheet1", "B1", 40d, "A1");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.Converged);
        Assert.Equal(20d, ReadDouble(batch, "A1"), 6);
        Assert.Equal(40d, ReadDouble(batch, "B1"), 6);
    }

    [Fact]
    public void GoalSeek_NullGoal_RejectsBeforeOpeningExcel()
    {
        var exception = Assert.Throws<ArgumentNullException>(
            () => _commands.GoalSeek(null!, "Sheet1", "B1", null, "A1"));

        Assert.Equal("goal", exception.ParamName);
    }

    [Fact]
    public void ScenarioLifecycle_CreateListShowUpdateDelete_ChangesRealCells()
    {
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);
        ConfigureSheet(batch, sheet =>
        {
            SetValue(sheet, "A1", 1d);
            SetValue(sheet, "A2", 2d);
        });

        var createResult = _commands.CreateScenario(
            batch,
            "Sheet1",
            "Best Case",
            "A1:A2",
            [10d, 20d],
            "Optimistic inputs",
            locked: false,
            hidden: false);
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var listResult = _commands.ListScenarios(batch, "Sheet1");
        var scenario = Assert.Single(listResult.Scenarios);
        Assert.Equal("Best Case", scenario.Name);
        Assert.Equal("$A$1:$A$2", scenario.ChangingCells);
        Assert.Equal([10d, 20d], scenario.Values.Select(Convert.ToDouble));

        var showResult = _commands.ShowScenario(batch, "Sheet1", "Best Case");
        Assert.True(showResult.Success, showResult.ErrorMessage);
        Assert.Equal(10d, ReadDouble(batch, "A1"), 6);
        Assert.Equal(20d, ReadDouble(batch, "A2"), 6);

        var updateResult = _commands.UpdateScenario(batch, "Sheet1", "Best Case", "A1:A2", [30d, 40d]);
        Assert.True(updateResult.Success, updateResult.ErrorMessage);
        _commands.ShowScenario(batch, "Sheet1", "Best Case");
        Assert.Equal(30d, ReadDouble(batch, "A1"), 6);
        Assert.Equal(40d, ReadDouble(batch, "A2"), 6);

        var deleteResult = _commands.DeleteScenario(batch, "Sheet1", "Best Case");
        Assert.True(deleteResult.Success, deleteResult.ErrorMessage);
        Assert.Empty(_commands.ListScenarios(batch, "Sheet1").Scenarios);
    }

    [Fact]
    public void ListScenarios_ReturnsMetadata()
    {
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);
        ConfigureSheet(batch, sheet => SetValue(sheet, "C1", 7d));
        _commands.CreateScenario(batch, "Sheet1", "Locked Plan", "C1", [11d], "Planning case", locked: true, hidden: true);

        var result = _commands.ListScenarios(batch, "Sheet1");

        Assert.True(result.Success, result.ErrorMessage);
        var scenario = Assert.Single(result.Scenarios);
        Assert.Contains("Planning case", scenario.Comment, StringComparison.Ordinal);
        Assert.True(scenario.Locked);
        Assert.True(scenario.Hidden);
    }

    [Theory]
    [InlineData(ScenarioSummaryType.Summary)]
    [InlineData(ScenarioSummaryType.PivotTable)]
    public void CreateScenarioSummary_AddsReportWorksheet(ScenarioSummaryType reportType)
    {
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);
        ConfigureSheet(batch, sheet =>
        {
            SetValue(sheet, "A1", 2d);
            SetFormula(sheet, "B1", "=A1*3");
        });
        _commands.CreateScenario(batch, "Sheet1", "Base", "A1", [2d]);
        _commands.CreateScenario(batch, "Sheet1", "Growth", "A1", [5d]);
        var sheetCountBefore = ReadWorksheetCount(batch);

        var result = _commands.CreateScenarioSummary(batch, "Sheet1", reportType, "B1");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(sheetCountBefore + 1, ReadWorksheetCount(batch));
        Assert.False(string.IsNullOrWhiteSpace(result.ReportSheetName));
        Assert.Equal(reportType == ScenarioSummaryType.Summary ? "summary" : "pivot-table", result.ReportType);
    }

    [Fact]
    public void CreateDataTable_OneVariable_PopulatesResults()
    {
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);
        ConfigureSheet(batch, sheet =>
        {
            SetValue(sheet, "D1", 0d);
            SetFormula(sheet, "B1", "=D1*2");
            SetValue(sheet, "A2", 1d);
            SetValue(sheet, "A3", 2d);
            SetValue(sheet, "A4", 3d);
        });

        var result = _commands.CreateDataTable(batch, "Sheet1", "A1:B4", columnInputCell: "D1");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2d, ReadDouble(batch, "B2"), 6);
        Assert.Equal(4d, ReadDouble(batch, "B3"), 6);
        Assert.Equal(6d, ReadDouble(batch, "B4"), 6);
    }

    [Fact]
    public void CreateDataTable_TwoVariable_PopulatesMatrix()
    {
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);
        ConfigureSheet(batch, sheet =>
        {
            SetValue(sheet, "A12", 0d);
            SetValue(sheet, "A13", 0d);
            SetFormula(sheet, "A1", "=A12*100+A13");
            SetValue(sheet, "B1", 2d);
            SetValue(sheet, "C1", 3d);
            SetValue(sheet, "A2", 4d);
            SetValue(sheet, "A3", 5d);
        });

        var result = _commands.CreateDataTable(batch, "Sheet1", "A1:C3", "A12", "A13");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(204d, ReadDouble(batch, "B2"), 6);
        Assert.Equal(305d, ReadDouble(batch, "C3"), 6);
    }

    private static void ConfigureSheet(IExcelBatch batch, Action<Excel.Worksheet> configure)
    {
        batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets["Sheet1"];
                configure(sheet);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static double ReadDouble(IExcelBatch batch, string address)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets["Sheet1"];
                range = sheet.Range[address];
                return Convert.ToDouble(range.Value2);
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static int ReadWorksheetCount(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                return sheets.Count;
            }
            finally
            {
                ComUtilities.Release(ref sheets);
            }
        });
    }

    private static void SetValue(Excel.Worksheet sheet, string address, object value)
    {
        Excel.Range? range = null;
        try
        {
            range = sheet.Range[address];
            range.Value2 = value;
        }
        finally
        {
            ComUtilities.Release(ref range);
        }
    }

    private static void SetFormula(Excel.Worksheet sheet, string address, string formula)
    {
        Excel.Range? range = null;
        try
        {
            range = sheet.Range[address];
            range.Formula = formula;
        }
        finally
        {
            ComUtilities.Release(ref range);
        }
    }
}
