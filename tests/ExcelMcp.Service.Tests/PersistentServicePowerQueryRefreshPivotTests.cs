using System.Globalization;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryRefreshTests
{
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
