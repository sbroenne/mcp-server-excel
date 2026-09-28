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
        var stopwatch = Stopwatch.StartNew();

        var result = _queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromSeconds(90));

        stopwatch.Stop();
        Assert.True(result.Success, $"Worksheet refresh failed: {result.ErrorMessage}");
        Assert.False(
            result.HasErrors,
            $"Refresh reported errors: {string.Join(", ", result.ErrorMessages)}");
        AssertCompletesWithoutDeadlock(stopwatch, "Worksheet refresh");
    }

    [Fact]
    public void Refresh_DataModelLoadedQuery_CompletesWithoutDeadlock()
    {
        var queryName = CreateDeadlockQuery(
            "DR_DataModel",
            PowerQueryLoadMode.LoadToDataModel);
        var stopwatch = Stopwatch.StartNew();

        var result = _queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromSeconds(90));

        stopwatch.Stop();
        Assert.True(result.Success, $"Data model refresh failed: {result.ErrorMessage}");
        Assert.False(
            result.HasErrors,
            $"Refresh reported errors: {string.Join(", ", result.ErrorMessages)}");
        AssertCompletesWithoutDeadlock(stopwatch, "Data model refresh");
    }

    [Fact]
    public void Evaluate_TemporaryWorksheetQuery_CompletesWithoutDeadlock()
    {
        var stopwatch = Stopwatch.StartNew();

        var result = _queries.Evaluate(
            _fixture.BatchToken,
            DeadlockRegressionMCode);

        stopwatch.Stop();
        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.Equal(3, result.RowCount);
        AssertCompletesWithoutDeadlock(stopwatch, "Evaluate");
    }

    [Fact]
    public void LoadTo_Table_CompletesWithoutDeadlock()
    {
        var queryName = CreateDeadlockQuery(
            "DR_LoadTable",
            PowerQueryLoadMode.ConnectionOnly);
        var sheetName = UniqueName("DR_Target");
        var stopwatch = Stopwatch.StartNew();

        var result = _queries.LoadTo(
            _fixture.BatchToken,
            queryName,
            PowerQueryLoadMode.LoadToTable,
            sheetName,
            "A1");
        _fixture.RegisterSheetForCleanup(sheetName);

        stopwatch.Stop();
        Assert.True(result.Success, $"LoadTo(Table) failed: {result.ErrorMessage}");
        AssertCompletesWithoutDeadlock(stopwatch, "LoadTo(Table)");
    }

    [Fact]
    public void LoadTo_DataModel_CompletesWithoutDeadlock()
    {
        var queryName = CreateDeadlockQuery(
            "DR_LoadModel",
            PowerQueryLoadMode.ConnectionOnly);
        var stopwatch = Stopwatch.StartNew();

        var result = _queries.LoadTo(
            _fixture.BatchToken,
            queryName,
            PowerQueryLoadMode.LoadToDataModel);

        stopwatch.Stop();
        Assert.True(result.Success, $"LoadTo(DataModel) failed: {result.ErrorMessage}");
        AssertCompletesWithoutDeadlock(stopwatch, "LoadTo(DataModel)");
    }

    [Fact]
    public void Update_WorksheetLoadedQuery_CompletesWithoutDeadlock()
    {
        var queryName = CreateDeadlockQuery(
            "DR_UpdateSheet",
            PowerQueryLoadMode.LoadToTable);
        var stopwatch = Stopwatch.StartNew();

        var result = _queries.Update(
            _fixture.BatchToken,
            queryName,
            UpdatedDeadlockMCode,
            refresh: true);

        stopwatch.Stop();
        Assert.True(result.Success, $"Update(worksheet) failed: {result.ErrorMessage}");
        AssertCompletesWithoutDeadlock(stopwatch, "Update(worksheet)");
    }

    [Fact]
    public void Update_DataModelLoadedQuery_CompletesWithoutDeadlock()
    {
        var queryName = CreateDeadlockQuery(
            "DR_UpdateModel",
            PowerQueryLoadMode.LoadToDataModel);
        var stopwatch = Stopwatch.StartNew();

        var result = _queries.Update(
            _fixture.BatchToken,
            queryName,
            UpdatedDeadlockMCode,
            refresh: true);

        stopwatch.Stop();
        Assert.True(result.Success, $"Update(data model) failed: {result.ErrorMessage}");
        AssertCompletesWithoutDeadlock(stopwatch, "Update(data model)");
    }

    private string CreateDeadlockQuery(
        string prefix,
        PowerQueryLoadMode loadMode)
    {
        var queryName = UniqueName(prefix);
        var targetSheet = loadMode == PowerQueryLoadMode.LoadToTable
            ? queryName
            : null;
        _queries.Create(
            _fixture.BatchToken,
            queryName,
            DeadlockRegressionMCode,
            loadMode,
            targetSheet);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        if (targetSheet is not null)
        {
            _fixture.RegisterSheetForCleanup(targetSheet);
        }

        return queryName;
    }

    private static void AssertCompletesWithoutDeadlock(
        Stopwatch stopwatch,
        string operation) =>
        Assert.True(
            stopwatch.Elapsed < TimeSpan.FromSeconds(60),
            $"{operation} took {stopwatch.Elapsed.TotalSeconds:F1}s - " +
            "possible COM deadlock regression.");
}
