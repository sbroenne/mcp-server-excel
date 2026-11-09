using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryRefreshTests
{
    private const string MultiPivotInitialCode = """
        let
            Source = #table(
                type table [Group = text, Amount = Int64.Type, Units = Int64.Type],
                {{"A", 10, 1}, {"B", 20, 2}, {"A", 30, 3}, {"B", 40, 4}})
        in
            Source
        """;
    private const string MultiPivotUpdatedCode = """
        let
            Source = #table(
                type table [Group = text, Amount = Int64.Type, Units = Int64.Type],
                {{"A", 100, 11}, {"B", 200, 22}, {"A", 300, 33}, {"B", 400, 44}})
        in
            Source
        """;

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task RefreshAll_RefreshesEachPivotTable(bool workbookRefresh, bool sharedCache)
    {
        var (queryName, reportSheet, amountPivot, unitsPivot) = CreateRefreshPivotPair(sharedCache);
        await _fixture.SaveAndReopenAsync();
        AssertPivotGroupTotals(amountPivot, 40, 60);
        AssertPivotGroupTotals(unitsPivot, 4, 6);
        RequireSuccess(_queries.Update(
            _fixture.BatchToken, queryName, MultiPivotUpdatedCode, refresh: false));
        AssertPivotGroupTotals(amountPivot, 40, 60);
        AssertPivotGroupTotals(unitsPivot, 4, 6);

        string[] expectedPivots = [$"{reportSheet}!{amountPivot}", $"{reportSheet}!{unitsPivot}"];
        if (workbookRefresh)
        {
            var result = RequireSuccess(_workbooks.RefreshAll(_fixture.BatchToken, TimeSpan.FromMinutes(1)));
            Assert.Equal([queryName], result.RefreshedQueries);
            Assert.Equal(expectedPivots.Order(), result.RefreshedPivotTables.Order());
            Assert.Empty(result.FailedQueries);
            Assert.Empty(result.FailedConnections);
            Assert.Empty(result.FailedPivotTables);
        }
        else
        {
            var result = RequireSuccess(_queries.RefreshAll(_fixture.BatchToken, TimeSpan.FromMinutes(1)));
            Assert.Equal([queryName], result.RefreshedQueries);
            Assert.Equal(expectedPivots.Order(), result.RefreshedPivotTables.Order());
            Assert.Empty(result.FailedQueries);
            Assert.Empty(result.FailedPivotTables);
        }

        PowerQueryStateAssertions.AssertStored(
            _fixture, queryName, MultiPivotUpdatedCode, PowerQueryLoadMode.LoadToTable, queryName,
            ["Group", "Amount", "Units"],
            [["A", 100, 11], ["B", 200, 22], ["A", 300, 33], ["B", 400, 44]]);
        AssertPivotGroupTotals(amountPivot, 400, 600);
        AssertPivotGroupTotals(unitsPivot, 44, 66);
    }

    private (string QueryName, string ReportSheet, string AmountPivot, string UnitsPivot)
        CreateRefreshPivotPair(bool sharedCache)
    {
        string queryName = UniqueName("MultiPivotQuery");
        string reportSheet = UniqueName("MultiPivotReport");
        string amountPivot = UniqueName("AmountPivot");
        string unitsPivot = UniqueName("UnitsPivot");
        RequireSuccess(_queries.Create(
            _fixture.BatchToken, queryName, MultiPivotInitialCode, PowerQueryLoadMode.LoadToTable, queryName));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);
        RequireSuccess(_sheets.Create(_fixture.BatchToken, reportSheet));
        _fixture.RegisterSheetForCleanup(reportSheet);

        RequireSuccess(_pivotTables.CreateFromTable(
            _fixture.BatchToken, queryName, reportSheet, "A3", amountPivot));
        RequireSuccess(_pivotFields.AddRowField(_fixture.BatchToken, amountPivot, "Group"));
        RequireSuccess(_pivotFields.AddValueField(
            _fixture.BatchToken, amountPivot, "Amount", AggregationFunction.Sum));
        RequireSuccess(_pivotTables.CreateFromTable(
            _fixture.BatchToken, queryName, reportSheet, "E3", unitsPivot,
            shareCacheWith: sharedCache ? amountPivot : null,
            shareCacheWithSheet: sharedCache ? reportSheet : null));
        RequireSuccess(_pivotFields.AddRowField(_fixture.BatchToken, unitsPivot, "Group"));
        RequireSuccess(_pivotFields.AddValueField(
            _fixture.BatchToken, unitsPivot, "Units", AggregationFunction.Sum));

        int amountCache = ReadPivotCacheIndex(reportSheet, amountPivot);
        int unitsCache = ReadPivotCacheIndex(reportSheet, unitsPivot);
        Assert.Equal(sharedCache, amountCache == unitsCache);
        return (queryName, reportSheet, amountPivot, unitsPivot);
    }

    private int ReadPivotCacheIndex(string sheetName, string pivotName) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? worksheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            try
            {
                worksheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(worksheet);
                pivots = worksheet.PivotTables();
                pivot = pivots.Item(pivotName);
                cache = pivot.PivotCache();
                return cache.Index;
            }
            finally
            {
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref worksheet);
            }
        });

    private void AssertPivotGroupTotals(string pivotName, double groupA, double groupB)
    {
        var calculations = _fixture.CreateCommands<IPivotTableCalcCommands>();
        var data = RequireSuccess(calculations.GetData(_fixture.BatchToken, pivotName));
        Assert.Equal(4, data.Values.Count);
        Assert.All(data.Values, row => Assert.Equal(2, row.Count));
        Assert.Equal("A", data.Values[1][0]);
        Assert.Equal(groupA, Convert.ToDouble(data.Values[1][1], CultureInfo.InvariantCulture));
        Assert.Equal("B", data.Values[2][0]);
        Assert.Equal(groupB, Convert.ToDouble(data.Values[2][1], CultureInfo.InvariantCulture));
        Assert.Equal(groupA + groupB, Convert.ToDouble(data.Values[3][1], CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RefreshAll_ReportsFailedPivotWithoutSkippingHealthyPivot(bool workbookRefresh)
    {
        var (queryName, reportSheet, amountPivot, unitsPivot) = CreateRefreshPivotPair(false);
        await _fixture.SaveAndReopenAsync();
        RequireSuccess(_pivotTables.SetCacheOptions(_fixture.BatchToken, unitsPivot, enableRefresh: false));
        AssertPivotGroupTotals(amountPivot, 40, 60);
        AssertPivotGroupTotals(unitsPivot, 4, 6);
        RequireSuccess(_queries.Update(_fixture.BatchToken, queryName, MultiPivotUpdatedCode, refresh: false));

        if (workbookRefresh)
        {
            var result = _workbooks.RefreshAll(_fixture.BatchToken, TimeSpan.FromMinutes(1));
            Assert.False(result.Success);
            Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
            Assert.Equal([queryName], result.RefreshedQueries);
            Assert.Equal([$"{reportSheet}!{amountPivot}"], result.RefreshedPivotTables);
            Assert.Empty(result.FailedQueries);
            Assert.Empty(result.FailedConnections);
            var failure = Assert.Single(result.FailedPivotTables);
            Assert.Equal(reportSheet, failure.SheetName);
            Assert.Equal(unitsPivot, failure.PivotTableName);
            Assert.False(string.IsNullOrWhiteSpace(failure.ErrorMessage));
            Assert.False(string.IsNullOrWhiteSpace(failure.ExceptionType));
            Assert.Matches("^0x[0-9A-F]{8}$", failure.HResult);
        }
        else
        {
            var result = _queries.RefreshAll(_fixture.BatchToken, TimeSpan.FromMinutes(1));
            Assert.False(result.Success);
            Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
            Assert.Equal([queryName], result.RefreshedQueries);
            Assert.Equal([$"{reportSheet}!{amountPivot}"], result.RefreshedPivotTables);
            Assert.Empty(result.FailedQueries);
            var failure = Assert.Single(result.FailedPivotTables);
            Assert.Equal($"{reportSheet}!{unitsPivot}", failure.PivotTableName);
            Assert.False(string.IsNullOrWhiteSpace(failure.ErrorMessage));
            Assert.False(string.IsNullOrWhiteSpace(failure.ExceptionType));
            Assert.Matches("^0x[0-9A-F]{8}$", failure.HResult);
        }

        PowerQueryStateAssertions.AssertStored(
            _fixture, queryName, MultiPivotUpdatedCode, PowerQueryLoadMode.LoadToTable, queryName,
            ["Group", "Amount", "Units"],
            [["A", 100, 11], ["B", 200, 22], ["A", 300, 33], ["B", 400, 44]]);
        AssertPivotGroupTotals(amountPivot, 400, 600);
        AssertPivotGroupTotals(unitsPivot, 4, 6);
    }

    [Fact]
    public void RefreshAll_RefreshesPivotTableBuiltOnRefreshedQueryByDefault()
    {
        const string initialCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 10}, {"B", 20}, {"A", 30}, {"B", 40}})
            in
                Source
            """;
        const string updatedCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 100}, {"B", 200}, {"A", 300}, {"B", 400}})
            in
                Source
            """;

        string queryName = UniqueName("RefreshPivotQuery");
        string reportSheet = UniqueName("RefreshPivotReport");
        string pivotName = UniqueName("RefreshPivotTable");
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            initialCode,
            PowerQueryLoadMode.LoadToTable,
            queryName));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);
        RequireSuccess(_sheets.Create(_fixture.BatchToken, reportSheet));
        _fixture.RegisterSheetForCleanup(reportSheet);

        RequireSuccess(_pivotTables.CreateFromTable(
            _fixture.BatchToken,
            queryName,
            reportSheet,
            "A3",
            pivotName));
        RequireSuccess(_pivotFields.AddRowField(_fixture.BatchToken, pivotName, "Group"));
        RequireSuccess(_pivotFields.AddValueField(
            _fixture.BatchToken,
            pivotName,
            "Amount",
            AggregationFunction.Sum));

        AssertPivotContainsNumber(reportSheet, 100);
        RequireSuccess(_queries.Update(
            _fixture.BatchToken,
            queryName,
            updatedCode,
            refresh: false));

        var result = RequireSuccess(_queries.RefreshAll(
            _fixture.BatchToken,
            TimeSpan.FromMinutes(1)));

        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
        Assert.Contains($"{reportSheet}!{pivotName}", result.RefreshedPivotTables);
        Assert.Empty(result.FailedQueries);
        Assert.Empty(result.FailedPivotTables);
        PowerQueryStateAssertions.AssertStored(
            _fixture,
            queryName,
            updatedCode,
            PowerQueryLoadMode.LoadToTable,
            queryName,
            ["Group", "Amount"],
            [["A", 100], ["B", 200], ["A", 300], ["B", 400]]);
        AssertPivotContainsNumber(reportSheet, 1000);
    }

    [Fact]
    public void RefreshAll_QueryOnlyOptionLeavesDependentPivotTableUnrefreshed()
    {
        const string initialCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 10}, {"B", 20}})
            in
                Source
            """;
        const string updatedCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 100}, {"B", 200}})
            in
                Source
            """;

        string queryName = UniqueName("RefreshPivotOptOutQuery");
        string reportSheet = UniqueName("RefreshPivotOptOutReport");
        string pivotName = UniqueName("RefreshPivotOptOutTable");
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            initialCode,
            PowerQueryLoadMode.LoadToTable,
            queryName));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);
        RequireSuccess(_sheets.Create(_fixture.BatchToken, reportSheet));
        _fixture.RegisterSheetForCleanup(reportSheet);
        RequireSuccess(_pivotTables.CreateFromTable(
            _fixture.BatchToken,
            queryName,
            reportSheet,
            "A3",
            pivotName));
        RequireSuccess(_pivotFields.AddRowField(_fixture.BatchToken, pivotName, "Group"));
        RequireSuccess(_pivotFields.AddValueField(
            _fixture.BatchToken,
            pivotName,
            "Amount",
            AggregationFunction.Sum));

        AssertPivotContainsNumber(reportSheet, 30);
        RequireSuccess(_queries.Update(
            _fixture.BatchToken,
            queryName,
            updatedCode,
            refresh: false));

        var result = RequireSuccess(_queries.RefreshAll(
            _fixture.BatchToken,
            TimeSpan.FromMinutes(1),
            refreshPivotTables: false));

        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
        Assert.Equal([queryName], result.RefreshedQueries);
        Assert.Empty(result.RefreshedPivotTables);
        Assert.Empty(result.FailedPivotTables);
        AssertPivotContainsNumber(reportSheet, 30);
    }

    [Fact]
    public void RefreshAll_RefreshesPivotTableUsingSubsetOfQueryTableRange()
    {
        const string initialCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 10}, {"B", 20}})
            in
                Source
            """;
        const string updatedCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 100}, {"B", 200}})
            in
                Source
            """;

        string queryName = UniqueName("RefreshPivotRangeQuery");
        string reportSheet = UniqueName("RefreshPivotRangeReport");
        string pivotName = UniqueName("RefreshPivotRangeTable");
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            initialCode,
            PowerQueryLoadMode.LoadToTable,
            queryName));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);
        RequireSuccess(_sheets.Create(_fixture.BatchToken, reportSheet));
        _fixture.RegisterSheetForCleanup(reportSheet);
        RequireSuccess(_pivotTables.CreateFromRange(
            _fixture.BatchToken,
            queryName,
            "A1:B2",
            reportSheet,
            "A3",
            pivotName));
        RequireSuccess(_pivotFields.AddRowField(_fixture.BatchToken, pivotName, "Group"));
        RequireSuccess(_pivotFields.AddValueField(
            _fixture.BatchToken,
            pivotName,
            "Amount",
            AggregationFunction.Sum));
        AssertPivotContainsNumber(reportSheet, 10);
        RequireSuccess(_queries.Update(
            _fixture.BatchToken,
            queryName,
            updatedCode,
            refresh: false));

        var result = RequireSuccess(_queries.RefreshAll(
            _fixture.BatchToken,
            TimeSpan.FromMinutes(1)));

        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
        Assert.Contains($"{reportSheet}!{pivotName}", result.RefreshedPivotTables);
        Assert.Empty(result.FailedPivotTables);
        AssertPivotContainsNumber(reportSheet, 100);
    }

    [Fact]
    public void RefreshAll_DoesNotRefreshRangePivotFromSimilarlyNamedWorksheet()
    {
        const string initialCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 10}, {"B", 20}})
            in
                Source
            """;
        const string updatedCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 100}, {"B", 200}})
            in
                Source
            """;

        string queryName = UniqueName("RefreshPivotSimilarSheetQuery");
        string querySheet = UniqueName("Sales Data");
        string unrelatedSheet = querySheet.Replace(" ", string.Empty, StringComparison.Ordinal);
        string reportSheet = UniqueName("RefreshPivotSimilarSheetReport");
        string pivotName = UniqueName("RefreshPivotSimilarSheet");
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            initialCode,
            PowerQueryLoadMode.LoadToTable,
            querySheet));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(querySheet);
        RequireSuccess(_sheets.Create(_fixture.BatchToken, unrelatedSheet));
        _fixture.RegisterSheetForCleanup(unrelatedSheet);
        RequireSuccess(_commands.SetValues(
            _fixture.BatchToken,
            unrelatedSheet,
            "A1:B3",
            [["Group", "Amount"], ["A", 10], ["B", 20]]));
        RequireSuccess(_sheets.Create(_fixture.BatchToken, reportSheet));
        _fixture.RegisterSheetForCleanup(reportSheet);
        RequireSuccess(_pivotTables.CreateFromRange(
            _fixture.BatchToken,
            unrelatedSheet,
            "A1:B3",
            reportSheet,
            "A3",
            pivotName));
        RequireSuccess(_pivotFields.AddRowField(_fixture.BatchToken, pivotName, "Group"));
        RequireSuccess(_pivotFields.AddValueField(
            _fixture.BatchToken,
            pivotName,
            "Amount",
            AggregationFunction.Sum));
        RequireSuccess(_queries.Update(
            _fixture.BatchToken,
            queryName,
            updatedCode,
            refresh: false));

        var result = RequireSuccess(_queries.RefreshAll(
            _fixture.BatchToken,
            TimeSpan.FromMinutes(1)));

        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
        Assert.Empty(result.RefreshedPivotTables);
        Assert.Empty(result.FailedPivotTables);
        AssertPivotContainsNumber(reportSheet, 30);
    }

    [Fact]
    public void WorkbookRefreshAll_RefreshesQueryAndPivotTable()
    {
        const string initialCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 10}, {"B", 20}})
            in
                Source
            """;
        const string updatedCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 100}, {"B", 200}})
            in
                Source
            """;

        string queryName = UniqueName("WorkbookRefreshQuery");
        string reportSheet = UniqueName("WorkbookRefreshReport");
        string pivotName = UniqueName("WorkbookRefreshPivot");
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            initialCode,
            PowerQueryLoadMode.LoadToTable,
            queryName));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);
        RequireSuccess(_sheets.Create(_fixture.BatchToken, reportSheet));
        _fixture.RegisterSheetForCleanup(reportSheet);
        RequireSuccess(_pivotTables.CreateFromTable(
            _fixture.BatchToken,
            queryName,
            reportSheet,
            "A3",
            pivotName));
        RequireSuccess(_pivotFields.AddRowField(_fixture.BatchToken, pivotName, "Group"));
        RequireSuccess(_pivotFields.AddValueField(
            _fixture.BatchToken,
            pivotName,
            "Amount",
            AggregationFunction.Sum));
        RequireSuccess(_queries.Update(
            _fixture.BatchToken,
            queryName,
            updatedCode,
            refresh: false));

        var result = RequireSuccess(_workbooks.RefreshAll(
            _fixture.BatchToken,
            TimeSpan.FromMinutes(1)));

        Assert.True(result.Success, $"Workbook refresh-all failed: {result.ErrorMessage}");
        Assert.Contains(queryName, result.RefreshedQueries);
        Assert.Contains($"{reportSheet}!{pivotName}", result.RefreshedPivotTables);
        Assert.Empty(result.FailedQueries);
        Assert.Empty(result.FailedConnections);
        Assert.Empty(result.FailedPivotTables);
        AssertPivotContainsNumber(reportSheet, 300);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WorkbookRefreshAll_RefreshesOrdinaryTextConnection(bool backgroundQuery)
    {
        string sheetName = UniqueName("WorkbookRefreshText");
        string queryTableName = UniqueName("WorkbookRefreshTextQuery");
        string sourcePath = _fixture.CreateInputFile(
            ".csv",
            "Name,Value\r\nAlpha,10\r\n");
        var queryTables = _fixture.CreateCommands<IQueryTableCommands>();

        RequireSuccess(_sheets.Create(_fixture.BatchToken, sheetName));
        _fixture.RegisterSheetForCleanup(sheetName);
        RequireSuccess(queryTables.CreateText(
            _fixture.BatchToken,
            queryTableName,
            sourcePath,
            sheetName,
            "A1"));

        string connectionName = _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? worksheet = null;
            Excel.QueryTables? worksheetQueryTables = null;
            Excel.QueryTable? queryTable = null;
            Excel.WorkbookConnection? connection = null;
            try
            {
                worksheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(worksheet);
                worksheetQueryTables = worksheet.QueryTables;
                queryTable = worksheetQueryTables.Item(queryTableName);
                queryTable.BackgroundQuery = backgroundQuery;
                connection = queryTable.WorkbookConnection;
                Assert.Equal(Excel.XlConnectionType.xlConnectionTypeTEXT, connection.Type);
                return connection.Name;
            }
            finally
            {
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref worksheetQueryTables);
                ComUtilities.Release(ref worksheet);
            }
        });
        _fixture.RegisterConnectionForCleanup(connectionName);
        File.WriteAllText(sourcePath, "Name,Value\r\nBeta,37\r\n");

        var attempted = _workbooks.RefreshAll(
            _fixture.BatchToken,
            TimeSpan.FromMinutes(1));
        Assert.True(attempted.Success, System.Text.Json.JsonSerializer.Serialize(attempted));
        var result = RequireSuccess(attempted);

        Assert.True(result.Success, $"Workbook refresh-all failed: {result.ErrorMessage}");
        Assert.Contains(connectionName, result.RefreshedConnections);
        Assert.Empty(result.FailedConnections);
        Assert.Empty(result.FailedQueries);
        Assert.Empty(result.FailedPivotTables);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? worksheet = null;
            Excel.QueryTables? queries = null;
            Excel.QueryTable? query = null;
            try
            {
                worksheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(worksheet);
                queries = worksheet.QueryTables;
                query = queries.Item(queryTableName);
                Assert.Equal(backgroundQuery, query.BackgroundQuery);
                Assert.False(query.Refreshing);
            }
            finally
            {
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queries);
                ComUtilities.Release(ref worksheet);
            }
        });
        var rows = RequireSuccess(_commands.GetValues(
            _fixture.BatchToken,
            sheetName,
            "A1:B2")).Values;
        Assert.Equal(["Name", "Value"], rows[0]);
        Assert.Equal("Beta", rows[1][0]);
        Assert.Equal(37, Convert.ToInt32(rows[1][1], CultureInfo.InvariantCulture));
    }

    [Fact]
    public void RefreshAll_RefreshesDataModelPivotTablesForRefreshedModelQuery()
    {
        const string initialCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 10}, {"B", 20}})
            in
                Source
            """;
        const string updatedCode = """
            let
                Source = #table(
                    type table [Group = text, Amount = Int64.Type],
                    {{"A", 100}, {"B", 200}})
            in
                Source
            """;

        string queryName = UniqueName("RefreshDataModelPivotQuery");
        string reportSheet = UniqueName("RefreshDataModelPivotReport");
        string pivotName = UniqueName("RefreshDataModelPivot");
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            initialCode,
            PowerQueryLoadMode.LoadToDataModel));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        RequireSuccess(_sheets.Create(_fixture.BatchToken, reportSheet));
        _fixture.RegisterSheetForCleanup(reportSheet);
        RequireSuccess(_pivotTables.CreateFromDataModel(
            _fixture.BatchToken,
            queryName,
            reportSheet,
            "A3",
            pivotName));
        RequireSuccess(_queries.Update(
            _fixture.BatchToken,
            queryName,
            updatedCode,
            refresh: false));

        var result = RequireSuccess(_queries.RefreshAll(
            _fixture.BatchToken,
            TimeSpan.FromMinutes(1)));

        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
        Assert.Contains(queryName, result.RefreshedQueries);
        Assert.Contains($"{reportSheet}!{pivotName}", result.RefreshedPivotTables);
        Assert.Empty(result.FailedPivotTables);
        PowerQueryStateAssertions.AssertStored(
            _fixture,
            queryName,
            updatedCode,
            PowerQueryLoadMode.LoadToDataModel,
            null,
            ["Group", "Amount"],
            [["A", 100], ["B", 200]]);
    }

    private void AssertPivotContainsNumber(string sheetName, double expected)
    {
        var result = RequireSuccess(_commands.GetValues(
            _fixture.BatchToken,
            sheetName,
            "A3:B8"));
        var numbers = result.Values
            .SelectMany(row => row)
            .Select(value => double.TryParse(
                Convert.ToString(value, CultureInfo.InvariantCulture),
                NumberStyles.Number,
                CultureInfo.InvariantCulture,
                out double number)
                    ? number
                    : (double?)null)
            .Where(number => number.HasValue)
            .Select(number => number!.Value);

        Assert.Contains(expected, numbers);
    }
}
