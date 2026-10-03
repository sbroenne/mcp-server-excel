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

        RequireSuccess(result);
        Assert.Equal("BugRegression_NonFirstRow", result.ChartName);
        Assert.Equal(ChartType.Line, result.ChartType);
        var charts = _chartCommands.List(batch);
        RequireSuccess(charts);
        Assert.Contains(
            charts.Charts,
            chart => chart.Name == "BugRegression_NonFirstRow");
        AssertRegressionSeries(result.ChartName, _sheetName);
    }

    [Fact]
    public void CreateFromRange_SheetNameWithSpaces_CreatesChart()
    {
        var batch = _fixture.BatchToken;
        const string sheetName = "Deal Summary";
        _fixture.CreateNamedTestSheet(batch, sheetName);
        RequireSuccess(_commands.SetValues(
            batch,
            sheetName,
            "A1:D6",
            [["Product", "Q1", "Q2", "Q3"],
             ["Widget A", 100, 150, 200],
             ["Widget B", 200, 250, 300],
             ["Widget C", 300, 350, 400],
             ["Widget D", 400, 450, 500],
             ["Widget E", 500, 550, 600]]));

        var result = _chartCommands.CreateFromRange(
            batch,
            sheetName,
            "A1:D6",
            ChartType.Line,
            50, 50, 400, 300,
            "BugRegression_SpacesInName");

        RequireSuccess(result);
        Assert.Equal("BugRegression_SpacesInName", result.ChartName);
        Assert.Equal(sheetName, result.SheetName);
        var read = RequireSuccess(_chartCommands.Read(batch, result.ChartName));
        Assert.Equal(["Q1", "Q2", "Q3"], read.Series.Select(series => series.Name));
        object[] categories = ["Widget A", "Widget B", "Widget C", "Widget D", "Widget E"];
        AssertSeriesData(result.ChartName, 1, "Q1", categories, [100, 200, 300, 400, 500], sheetName);
        AssertSeriesData(result.ChartName, 2, "Q2", categories, [150, 250, 350, 450, 550], sheetName);
        AssertSeriesData(result.ChartName, 3, "Q3", categories, [200, 300, 400, 500, 600], sheetName);
    }

    [Fact]
    public void CreateFromRange_SheetWithSpacesAndDataAtNonFirstRow_CreatesChart()
    {
        var batch = _fixture.BatchToken;
        const string sheetName = "Export Data";
        _fixture.CreateNamedTestSheet(batch, sheetName);
        RequireSuccess(_commands.SetValues(
            batch,
            sheetName,
            "A9:D14",
            [["Service", "Current", "Proposed", "Delta"],
             ["Compute", 50000, 45000, -5000],
             ["Storage", 20000, 18000, -2000],
             ["Network", 15000, 14000, -1000],
             ["Database", 30000, 25000, -5000],
             ["AI/ML", 10000, 12000, 2000]]));

        var result = _chartCommands.CreateFromRange(
            batch,
            sheetName,
            "A9:D14",
            ChartType.Line,
            50, 50, 400, 300,
            "BugRegression_Combined");

        RequireSuccess(result);
        Assert.Equal("BugRegression_Combined", result.ChartName);
        Assert.Equal(sheetName, result.SheetName);
        Assert.Equal(ChartType.Line, result.ChartType);
        var read = RequireSuccess(_chartCommands.Read(batch, result.ChartName));
        Assert.Equal(["Current", "Proposed", "Delta"], read.Series.Select(series => series.Name));
        object[] categories = ["Compute", "Storage", "Network", "Database", "AI/ML"];
        AssertSeriesData(result.ChartName, 1, "Current", categories, [50000, 20000, 15000, 30000, 10000], sheetName);
        AssertSeriesData(result.ChartName, 2, "Proposed", categories, [45000, 18000, 14000, 25000, 12000], sheetName);
        AssertSeriesData(result.ChartName, 3, "Delta", categories, [-5000, -2000, -1000, -5000, 2000], sheetName);
    }

    [Fact]
    public void CreateFromTable_DataAtNonFirstRow_SucceedsAsWorkaround()
    {
        var batch = _fixture.BatchToken;
        SetRegressionData(batch, _sheetName);
        RequireSuccess(_tableCommands.Create(
            batch,
            _sheetName,
            "BugWorkaroundTable",
            "A9:D14",
            true));

        var result = _chartCommands.CreateFromTable(
            batch,
            "BugWorkaroundTable",
            _sheetName,
            ChartType.Line,
            50, 50, 400, 300,
            "BugWorkaround_Table");

        RequireSuccess(result);
        Assert.Equal("BugWorkaround_Table", result.ChartName);
        Assert.Equal(ChartType.Line, result.ChartType);
        AssertRegressionSeries(result.ChartName, _sheetName);
    }

    private void SetRegressionData(
        Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch,
        string sheetName) =>
        RequireSuccess(_commands.SetValues(
            batch,
            sheetName,
            "A9:D14",
            [["Quarter", "Revenue", "Cost", "Profit"],
             ["Q1", 1200, 800, 400],
             ["Q2", 1500, 900, 600],
             ["Q3", 1800, 1000, 800],
             ["Q4", 2100, 1100, 1000],
             ["Q5", 2400, 1200, 1200]]));

    private void AssertRegressionSeries(string chartName, string sheetName)
    {
        var read = RequireSuccess(_chartCommands.Read(_fixture.BatchToken, chartName));
        Assert.Equal(["Revenue", "Cost", "Profit"], read.Series.Select(series => series.Name));
        object[] categories = ["Q1", "Q2", "Q3", "Q4", "Q5"];
        AssertSeriesData(chartName, 1, "Revenue", categories, [1200, 1500, 1800, 2100, 2400], sheetName);
        AssertSeriesData(chartName, 2, "Cost", categories, [800, 900, 1000, 1100, 1200], sheetName);
        AssertSeriesData(chartName, 3, "Profit", categories, [400, 600, 800, 1000, 1200], sheetName);
    }
}
