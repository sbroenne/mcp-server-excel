using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceChartFormattingTests
{
    [Fact]
    public void CreateFromRange_DataAtNonFirstRow_CreatesChart()
    {
        var batch = _fixture.BatchToken;
        SetRegressionData(batch, _sheetName);

        var result = _chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A9:D14",
            ChartType.Line,
            50, 50, 400, 300,
            "BugRegression_NonFirstRow");

        Assert.True(result.Success, "CreateFromRange failed: chart was not created");
        Assert.Equal("BugRegression_NonFirstRow", result.ChartName);
        Assert.Equal(ChartType.Line, result.ChartType);
        var charts = _chartCommands.List(batch);
        Assert.True(charts.Success);
        Assert.Contains(
            charts.Charts,
            chart => chart.Name == "BugRegression_NonFirstRow");
    }

    [Fact]
    public void CreateFromRange_SheetNameWithSpaces_CreatesChart()
    {
        var batch = _fixture.BatchToken;
        const string sheetName = "Deal Summary";
        _fixture.CreateNamedTestSheet(batch, sheetName);
        _commands.SetValues(
            batch,
            sheetName,
            "A1:D6",
            [["Product", "Q1", "Q2", "Q3"],
             ["Widget A", 100, 150, 200],
             ["Widget B", 200, 250, 300],
             ["Widget C", 300, 350, 400],
             ["Widget D", 400, 450, 500],
             ["Widget E", 500, 550, 600]]);

        var result = _chartCommands.CreateFromRange(
            batch,
            sheetName,
            "A1:D6",
            ChartType.Line,
            50, 50, 400, 300,
            "BugRegression_SpacesInName");

        Assert.True(result.Success, "CreateFromRange failed for sheet with spaces");
        Assert.Equal("BugRegression_SpacesInName", result.ChartName);
        Assert.Equal(sheetName, result.SheetName);
    }

    [Fact]
    public void CreateFromRange_SheetWithSpacesAndDataAtNonFirstRow_CreatesChart()
    {
        var batch = _fixture.BatchToken;
        const string sheetName = "Export Data";
        _fixture.CreateNamedTestSheet(batch, sheetName);
        _commands.SetValues(
            batch,
            sheetName,
            "A9:D14",
            [["Service", "Current", "Proposed", "Delta"],
             ["Compute", 50000, 45000, -5000],
             ["Storage", 20000, 18000, -2000],
             ["Network", 15000, 14000, -1000],
             ["Database", 30000, 25000, -5000],
             ["AI/ML", 10000, 12000, 2000]]);

        var result = _chartCommands.CreateFromRange(
            batch,
            sheetName,
            "A9:D14",
            ChartType.Line,
            50, 50, 400, 300,
            "BugRegression_Combined");

        Assert.True(result.Success, "CreateFromRange failed for combined scenario");
        Assert.Equal("BugRegression_Combined", result.ChartName);
        Assert.Equal(sheetName, result.SheetName);
        Assert.Equal(ChartType.Line, result.ChartType);
    }

    [Fact]
    public void CreateFromTable_DataAtNonFirstRow_SucceedsAsWorkaround()
    {
        var batch = _fixture.BatchToken;
        SetRegressionData(batch, _sheetName);
        _tableCommands.Create(
            batch,
            _sheetName,
            "BugWorkaroundTable",
            "A9:D14",
            true);

        var result = _chartCommands.CreateFromTable(
            batch,
            "BugWorkaroundTable",
            _sheetName,
            ChartType.Line,
            50, 50, 400, 300,
            "BugWorkaround_Table");

        Assert.True(result.Success, "CreateFromTable workaround failed unexpectedly");
        Assert.Equal("BugWorkaround_Table", result.ChartName);
        Assert.Equal(ChartType.Line, result.ChartType);
    }

    private void SetRegressionData(
        Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch,
        string sheetName) =>
        _commands.SetValues(
            batch,
            sheetName,
            "A9:D14",
            [["Quarter", "Revenue", "Cost", "Profit"],
             ["Q1", 1200, 800, 400],
             ["Q2", 1500, 900, 600],
             ["Q3", 1800, 1000, 800],
             ["Q4", 2100, 1100, 1000],
             ["Q5", 2400, 1200, 1200]]);
}
