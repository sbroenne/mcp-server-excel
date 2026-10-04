using System.Diagnostics;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryRefreshTests
{
    private const string DeadlockRegressionMCode = """
        let
            Source = #table(
                {"ID", "Name", "Value"},
                {
                    {1, "Alpha", 100},
                    {2, "Beta", 200},
                    {3, "Gamma", 300}
                })
        in
            Source
        """;

    private const string UpdatedDeadlockMCode = """
        let
            Source = #table(
                {"ID", "Name", "Value", "Extra"},
                {
                    {1, "Alpha", 100, "A"},
                    {2, "Beta", 200, "B"},
                    {3, "Gamma", 300, "C"}
                })
        in
            Source
        """;

    [Fact]
    public void Refresh_WorksheetLoadedQuery_CompletesWithoutDeadlock()
    {
        var queryName = CreateDeadlockQuery(
            "DR_Worksheet",
            PowerQueryLoadMode.LoadToTable);
        StageDeadlockQueryUpdate(queryName, PowerQueryLoadMode.LoadToTable);
        var stopwatch = Stopwatch.StartNew();

        var result = RequireSuccess(_queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromSeconds(90)));

        stopwatch.Stop();
        Assert.True(result.Success, $"Worksheet refresh failed: {result.ErrorMessage}");
        Assert.False(
            result.HasErrors,
            $"Refresh reported errors: {string.Join(", ", result.ErrorMessages)}");
        AssertCompletesWithoutDeadlock(stopwatch, "Worksheet refresh");
        AssertRefreshMetadata(result, queryName, queryName);
        AssertDeadlockQueryData(queryName, PowerQueryLoadMode.LoadToTable, updated: true);
    }

    [Fact]
    public void Refresh_DataModelLoadedQuery_CompletesWithoutDeadlock()
    {
        var queryName = CreateDeadlockQuery(
            "DR_DataModel",
            PowerQueryLoadMode.LoadToDataModel);
        StageDeadlockQueryUpdate(queryName, PowerQueryLoadMode.LoadToDataModel);
        var stopwatch = Stopwatch.StartNew();

        var result = RequireSuccess(_queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromSeconds(90)));

        stopwatch.Stop();
        Assert.True(result.Success, $"Data model refresh failed: {result.ErrorMessage}");
        Assert.False(
            result.HasErrors,
            $"Refresh reported errors: {string.Join(", ", result.ErrorMessages)}");
        AssertCompletesWithoutDeadlock(stopwatch, "Data model refresh");
        AssertRefreshMetadata(result, queryName, null);
        AssertDeadlockQueryData(queryName, PowerQueryLoadMode.LoadToDataModel, updated: true);
    }

    [Fact]
    public void Evaluate_TemporaryWorksheetQuery_CompletesWithoutDeadlock()
    {
        var guard = CreateRefreshGuard();
        var before = SnapshotQueries();
        var stopwatch = Stopwatch.StartNew();

        var result = RequireSuccess(_queries.Evaluate(
            _fixture.BatchToken,
            DeadlockRegressionMCode));

        stopwatch.Stop();
        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.Equal(3, result.RowCount);
        Assert.Equal(3, result.Columns.Count);
        Assert.Equal(["ID", "Name", "Value"], result.Columns);
        Assert.Equal(3, result.Rows.Count);
        Assert.Equal(DeadlockRegressionMCode, result.MCode);
        Assert.Equal(3, result.ColumnCount);
        PowerQueryStateAssertions.AssertRows([[1, "Alpha", 100], [2, "Beta", 200], [3, "Gamma", 300]],
            result.Rows);
        AssertCompletesWithoutDeadlock(stopwatch, "Evaluate");
        Assert.Equal(before, SnapshotQueries());
        AssertRefreshGuard(guard);
    }

    [Fact]
    public void LoadTo_Table_CompletesWithoutDeadlock()
    {
        var queryName = CreateDeadlockQuery(
            "DR_LoadTable",
            PowerQueryLoadMode.ConnectionOnly);
        var sheetName = UniqueName("DR_Target");
        var stopwatch = Stopwatch.StartNew();

        var result = RequireSuccess(_queries.LoadTo(
            _fixture.BatchToken,
            queryName,
            PowerQueryLoadMode.LoadToTable,
            sheetName,
            "A1"));
        _fixture.RegisterSheetForCleanup(sheetName);

        stopwatch.Stop();
        Assert.True(result.Success, $"LoadTo(Table) failed: {result.ErrorMessage}");
        AssertCompletesWithoutDeadlock(stopwatch, "LoadTo(Table)");
        AssertDeadlockQueryData(queryName, PowerQueryLoadMode.LoadToTable, sheetName: sheetName);
    }

    [Fact]
    public void LoadTo_DataModel_CompletesWithoutDeadlock()
    {
        var queryName = CreateDeadlockQuery(
            "DR_LoadModel",
            PowerQueryLoadMode.ConnectionOnly);
        var stopwatch = Stopwatch.StartNew();

        var result = RequireSuccess(_queries.LoadTo(
            _fixture.BatchToken,
            queryName,
            PowerQueryLoadMode.LoadToDataModel));

        stopwatch.Stop();
        Assert.True(result.Success, $"LoadTo(DataModel) failed: {result.ErrorMessage}");
        AssertCompletesWithoutDeadlock(stopwatch, "LoadTo(DataModel)");
        AssertDeadlockQueryData(queryName, PowerQueryLoadMode.LoadToDataModel);
    }

    [Fact]
    public void Update_WorksheetLoadedQuery_CompletesWithoutDeadlock()
    {
        var queryName = CreateDeadlockQuery(
            "DR_UpdateSheet",
            PowerQueryLoadMode.LoadToTable);
        var stopwatch = Stopwatch.StartNew();

        var result = RequireSuccess(_queries.Update(
            _fixture.BatchToken,
            queryName,
            UpdatedDeadlockMCode,
            refresh: true));
        _storedSources[queryName] = UpdatedDeadlockMCode;

        stopwatch.Stop();
        Assert.True(result.Success, $"Update(worksheet) failed: {result.ErrorMessage}");
        AssertCompletesWithoutDeadlock(stopwatch, "Update(worksheet)");
        AssertDeadlockQueryData(queryName, PowerQueryLoadMode.LoadToTable, updated: true);
    }

    [Fact]
    public void Update_DataModelLoadedQuery_CompletesWithoutDeadlock()
    {
        var queryName = CreateDeadlockQuery(
            "DR_UpdateModel",
            PowerQueryLoadMode.LoadToDataModel);
        var stopwatch = Stopwatch.StartNew();

        var result = RequireSuccess(_queries.Update(
            _fixture.BatchToken,
            queryName,
            UpdatedDeadlockMCode,
            refresh: true));
        _storedSources[queryName] = UpdatedDeadlockMCode;

        stopwatch.Stop();
        Assert.True(result.Success, $"Update(data model) failed: {result.ErrorMessage}");
        AssertCompletesWithoutDeadlock(stopwatch, "Update(data model)");
        AssertDeadlockQueryData(queryName, PowerQueryLoadMode.LoadToDataModel, updated: true);
    }

    private string CreateDeadlockQuery(
        string prefix,
        PowerQueryLoadMode loadMode)
    {
        var queryName = UniqueName(prefix);
        var targetSheet = loadMode == PowerQueryLoadMode.LoadToTable
            ? queryName
            : null;
        var created = RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            DeadlockRegressionMCode,
            loadMode,
            targetSheet));
        Assert.True(created.Success, created.ErrorMessage);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _storedSources.Add(queryName, DeadlockRegressionMCode);
        if (targetSheet is not null)
        {
            _fixture.RegisterSheetForCleanup(targetSheet);
        }

        PowerQueryStateAssertions.AssertStored(_fixture, queryName, DeadlockRegressionMCode,
            loadMode, targetSheet, ["ID", "Name", "Value"],
            [[1, "Alpha", 100], [2, "Beta", 200], [3, "Gamma", 300]]);
        return queryName;
    }

    private void StageDeadlockQueryUpdate(string queryName, PowerQueryLoadMode loadMode)
    {
        AssertDeadlockQueryData(queryName, loadMode);
        StageSource(queryName, UpdatedDeadlockMCode);
        AssertDeadlockQueryData(queryName, loadMode);
    }

    private void AssertDeadlockQueryData(string queryName, PowerQueryLoadMode loadMode,
        bool updated = false, string? sheetName = null)
    {
        string[] columns = updated ? ["ID", "Name", "Value", "Extra"] : ["ID", "Name", "Value"];
        object[][] rows = updated
            ? [[1, "Alpha", 100, "A"], [2, "Beta", 200, "B"], [3, "Gamma", 300, "C"]]
            : [[1, "Alpha", 100], [2, "Beta", 200], [3, "Gamma", 300]];
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, _storedSources[queryName],
            loadMode, loadMode == PowerQueryLoadMode.LoadToTable ? sheetName ?? queryName : null,
            columns, rows);
    }

    private static void AssertCompletesWithoutDeadlock(
        Stopwatch stopwatch,
        string operation) =>
        Assert.True(
            stopwatch.Elapsed < TimeSpan.FromSeconds(60),
            $"{operation} took {stopwatch.Elapsed.TotalSeconds:F1}s - " +
            "possible COM deadlock regression.");
}
