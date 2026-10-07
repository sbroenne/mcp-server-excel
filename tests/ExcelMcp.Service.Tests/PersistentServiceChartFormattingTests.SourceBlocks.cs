using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceChartFormattingTests
{
    private static readonly object[] SeedCategories = ["A", "B", "C", "D", "E"];

    [Fact]
    public void CreateFromRange_SeparateBlocks_UsesFirstBlockAsLabels()
    {
        var batch = _fixture.BatchToken;
        var before = CountCharts(_sheetName);

        var result = RequireSuccess(_chartCommands.CreateFromRange(
            batch, _sheetName, "A1:A6,C1:C6", ChartType.ColumnClustered, 50, 50));

        Assert.Equal(_sheetName, result.SheetName);
        Assert.Equal(before + 1, CountCharts(_sheetName));
        var read = RequireSuccess(_chartCommands.Read(batch, result.ChartName));
        Assert.Equal(["Series2"], read.Series.Select(series => series.Name));
        AssertSeriesData(result.ChartName, 1, "Series2", SeedCategories, [20, 25, 30, 35, 40]);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CreateFromRange_SeparateBlocksOnQuotedSheet_ReadsThatSheet(bool qualifyEveryBlock)
    {
        var batch = _fixture.BatchToken;
        var sourceSheet = CreateSourceSheet("Bob's Data");
        var prefix = QuoteSheet(sourceSheet);
        var source = qualifyEveryBlock
            ? $"{prefix}A1:A6, {prefix}C1:C6"
            : $"{prefix}A1:A6,C1:C6";

        var result = RequireSuccess(_chartCommands.CreateFromRange(
            batch, _sheetName, source, ChartType.ColumnClustered, 50, 50));

        Assert.Equal(_sheetName, result.SheetName);
        AssertSeriesData(result.ChartName, 1, "Other2", SeedCategories, [200, 250, 300, 350, 400]);
        Assert.Single(RequireSuccess(_chartCommands.Read(batch, result.ChartName)).Series);
    }

    [Fact]
    public void CreateFromRange_BlocksOnDifferentSheets_RejectsBeforeCreatingChart()
    {
        var batch = _fixture.BatchToken;
        var sourceSheet = CreateSourceSheet("Split Source");
        var before = CountCharts(_sheetName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.CreateFromRange(
                batch,
                _sheetName,
                $"{QuoteSheet(sourceSheet)}A1:A6,{QuoteSheet(_sheetName)}C1:C6",
                ChartType.ColumnClustered));

        Assert.Contains("same sheet", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(sourceSheet, exception.Message, StringComparison.Ordinal);
        Assert.Contains(_sheetName, exception.Message, StringComparison.Ordinal);
        Assert.Equal(before, CountCharts(_sheetName));
    }

    [Theory]
    [InlineData("A1:A6,not-a-range")]
    [InlineData("A1:A6,,C1:C6")]
    [InlineData("'Unclosed!A1:A6")]
    public void CreateFromRange_InvalidSourceAddress_RejectsBeforeCreatingChart(string source)
    {
        var batch = _fixture.BatchToken;
        var before = CountCharts(_sheetName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.CreateFromRange(batch, _sheetName, source, ChartType.ColumnClustered));

        Assert.Contains(source, exception.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("0x800A03EC", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, CountCharts(_sheetName));
    }

    [Fact]
    public void CreateFromRange_MissingSourceSheet_RejectsBeforeCreatingChart()
    {
        var batch = _fixture.BatchToken;
        var before = CountCharts(_sheetName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.CreateFromRange(
                batch, _sheetName, "'No Such Sheet'!A1:B4", ChartType.ColumnClustered));

        Assert.Contains("No Such Sheet", exception.Message, StringComparison.Ordinal);
        Assert.Equal(before, CountCharts(_sheetName));
    }

    [Fact]
    public void CreateFromRange_NameRejectedAfterChartExists_ReportsLeftoverChart()
    {
        var batch = _fixture.BatchToken;
        var namesBefore = ChartNames(_sheetName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.CreateFromRange(
                batch, _sheetName, "A1:B4", ChartType.Line, 50, 50,
                chartName: new string('x', 300)));

        var leftover = Assert.Single(ChartNames(_sheetName).Except(namesBefore));
        LeftoverObjectAssert.Reported(exception.Message, "chart", leftover, _sheetName);
        AssertSeriesData(leftover, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    [Fact]
    public void CreateFromTable_NameRejectedAfterChartExists_ReportsLeftoverChart()
    {
        var batch = _fixture.BatchToken;
        const string tableName = "LeftoverChartTable";
        RequireSuccess(_tableCommands.Create(batch, _sheetName, tableName, "A1:C6", true));
        var namesBefore = ChartNames(_sheetName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.CreateFromTable(
                batch, tableName, _sheetName, ChartType.Line, 50, 50,
                chartName: new string('x', 300)));

        var leftover = Assert.Single(ChartNames(_sheetName).Except(namesBefore));
        LeftoverObjectAssert.Reported(exception.Message, "chart", leftover, _sheetName);
    }

    [Fact]
    public void SetSourceRange_UnqualifiedWhileAnotherSheetIsActive_UsesChartSheet()
    {
        var batch = _fixture.BatchToken;
        var created = RequireSuccess(_chartCommands.CreateFromRange(
            batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));
        var decoySheet = CreateSourceSheet("Decoy");
        ActivateSheet(decoySheet);

        RequireSuccess(_chartCommands.SetSourceRange(batch, created.ChartName, "A1:C6"));

        var read = RequireSuccess(_chartCommands.Read(batch, created.ChartName));
        Assert.Equal(["Series1", "Series2"], read.Series.Select(series => series.Name));
        AssertSeriesData(created.ChartName, 1, "Series1", SeedCategories, [10, 15, 20, 25, 30]);
        AssertSeriesData(created.ChartName, 2, "Series2", SeedCategories, [20, 25, 30, 35, 40]);
    }

    [Fact]
    public void SetSourceRange_SeparateBlocksOnQuotedSheet_UpdatesChart()
    {
        var batch = _fixture.BatchToken;
        var created = RequireSuccess(_chartCommands.CreateFromRange(
            batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));
        var sourceSheet = CreateSourceSheet("Bob's Data");
        ActivateSheet(_sheetName);
        var prefix = QuoteSheet(sourceSheet);

        RequireSuccess(_chartCommands.SetSourceRange(
            batch, created.ChartName, $"{prefix}A1:A6,{prefix}C1:C6"));

        Assert.Single(RequireSuccess(_chartCommands.Read(batch, created.ChartName)).Series);
        AssertSeriesData(created.ChartName, 1, "Other2", SeedCategories, [200, 250, 300, 350, 400]);
    }

    [Fact]
    public void SetSourceRange_UnqualifiedSeparateBlocks_UsesChartSheet()
    {
        var batch = _fixture.BatchToken;
        var created = RequireSuccess(_chartCommands.CreateFromRange(
            batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));
        ActivateSheet(CreateSourceSheet("Decoy"));

        RequireSuccess(_chartCommands.SetSourceRange(batch, created.ChartName, "A1:A6,C1:C6"));

        Assert.Single(RequireSuccess(_chartCommands.Read(batch, created.ChartName)).Series);
        AssertSeriesData(created.ChartName, 1, "Series2", SeedCategories, [20, 25, 30, 35, 40]);
    }

    [Fact]
    public void SetSourceRange_BlocksOnDifferentSheets_LeavesChartUnchanged()
    {
        var batch = _fixture.BatchToken;
        var created = RequireSuccess(_chartCommands.CreateFromRange(
            batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));
        var before = RequireSuccess(_chartCommands.Read(batch, created.ChartName));
        var sourceSheet = CreateSourceSheet("Split Source");

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.SetSourceRange(
                batch,
                created.ChartName,
                $"{QuoteSheet(sourceSheet)}A1:A6,{QuoteSheet(_sheetName)}C1:C6"));

        Assert.Contains("same sheet", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertChartUnchanged(before);
        AssertSeriesData(created.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CreateFromRange_SheetNameInDifferentCase_UsesExistingSheet(bool qualifiedSource)
    {
        var batch = _fixture.BatchToken;
        var sourceSheet = CreateSourceSheet("Case Source");
        var sheetArgument = qualifiedSource ? _sheetName : _sheetName.ToUpperInvariant();
        var source = qualifiedSource
            ? $"{QuoteSheet(sourceSheet.ToUpperInvariant())}A1:A6,C1:C6"
            : "A1:A6,C1:C6";

        var result = RequireSuccess(_chartCommands.CreateFromRange(
            batch, sheetArgument, source, ChartType.ColumnClustered, 50, 50));

        var expectedName = qualifiedSource ? "Other2" : "Series2";
        double[] expectedValues = qualifiedSource ? [200, 250, 300, 350, 400] : [20, 25, 30, 35, 40];
        AssertSeriesData(result.ChartName, 1, expectedName, SeedCategories, expectedValues);
    }

    [Fact]
    public void SetSourceRange_SheetNameInDifferentCase_UsesExistingSheet()
    {
        var batch = _fixture.BatchToken;
        var created = RequireSuccess(_chartCommands.CreateFromRange(
            batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));
        var sourceSheet = CreateSourceSheet("Case Source");

        RequireSuccess(_chartCommands.SetSourceRange(
            batch, created.ChartName, $"{QuoteSheet(sourceSheet.ToUpperInvariant())}A1:A6,C1:C6"));

        AssertSeriesData(created.ChartName, 1, "Other2", SeedCategories, [200, 250, 300, 350, 400]);
    }

    private string CreateSourceSheet(string prefix)
    {
        var batch = _fixture.BatchToken;
        var sheetName = $"{prefix} {Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);
        RequireSuccess(_commands.SetValues(
            batch,
            sheetName,
            "A1:C6",
            [
                ["Category", "Other1", "Other2"],
                ["A", 100, 200],
                ["B", 150, 250],
                ["C", 200, 300],
                ["D", 250, 350],
                ["E", 300, 400],
            ]));
        return sheetName;
    }

    private static string QuoteSheet(string sheetName) =>
        $"'{sheetName.Replace("'", "''", StringComparison.Ordinal)}'!";

    private int CountCharts(string sheetName) => ChartNames(sheetName).Count;

    private List<string> ChartNames(string sheetName) =>
        [.. RequireSuccess(_chartCommands.List(_fixture.BatchToken)).Charts
            .Where(chart => string.Equals(chart.SheetName, sheetName, StringComparison.Ordinal))
            .Select(chart => chart.Name)];

    private void ActivateSheet(string sheetName) =>
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                sheet.Activate();
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
}
