using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServicePowerQueryLifecycleTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private const string TableMCode = """
        let
            Source = #table(
                {"Column1", "Column2", "Column3"},
                {
                    {"Value1", "Value2", "Value3"},
                    {"A", "B", "C"},
                    {"X", "Y", "Z"}
                }
            )
        in
            Source
        """;

    private readonly IPowerQueryCommands _queries =
        ServiceCommandProxy.Create<IPowerQueryCommands>(fixture);
    private readonly IDataModelCommands _dataModel =
        ServiceCommandProxy.Create<IDataModelCommands>(fixture);
    private readonly IConnectionCommands _connections =
        ServiceCommandProxy.Create<IConnectionCommands>(fixture);

    [Fact]
    public void Import_ValidMCode_ReturnsSuccess()
    {
        const string queryName = "TestQuery";
        var batch = _fixture.BatchToken;

        RequireSuccess(_queries.Create(batch, queryName, TableMCode));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);

        var result = RequireSuccess(_queries.List(batch));
        Assert.Contains(result.Queries, query => query.Name == queryName);
        AssertInitialTable(queryName);
    }

    [Fact]
    public void Update_ExistingQuery_ReturnsSuccess()
    {
        var queryName = "PQ_Update_" + Guid.NewGuid().ToString("N")[..8];
        const string updatedMCode = """
            let
                UpdatedSource = #table({"Column1", "Column2", "Column3"},
                    {{"Changed1", "Changed2", "Changed3"}, {"D", "E", "F"}})
            in
                UpdatedSource
            """;
        var batch = _fixture.BatchToken;
        RequireSuccess(_queries.Create(batch, queryName, TableMCode));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);
        AssertInitialTable(queryName);

        RequireSuccess(_queries.Update(batch, queryName, updatedMCode));
        var result = RequireSuccess(_queries.View(batch, queryName));
        Assert.Equal(updatedMCode, result.MCode);
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, updatedMCode,
            PowerQueryLoadMode.LoadToTable, queryName, ["Column1", "Column2", "Column3"],
            [["Changed1", "Changed2", "Changed3"], ["D", "E", "F"]]);
        Assert.Null(RequireSuccess(_commands.GetValues(batch, queryName, "A4")).Values[0][0]);
    }

    [Fact]
    public void Update_ExistingQuery_ReplacesNotMergesMCode()
    {
        var queryName = "PQ_ReplaceTest_" + Guid.NewGuid().ToString("N")[..8];
        const string originalMCode = """
            let
                OriginalSource = #table({"Marker"}, {{"ORIGINAL_MARKER"}}),
                OriginalStep = "Should be completely removed"
            in
                OriginalSource
            """;
        const string newMCode = """
            let
                NewSource = #table({"Marker"}, {{"NEW_MARKER"}}),
                NewStep = "Should be the only content"
            in
                NewSource
            """;
        var batch = _fixture.BatchToken;
        RequireSuccess(_queries.Create(batch, queryName, originalMCode,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        AssertStoredAndEvaluated(queryName, originalMCode, ["Marker"], [["ORIGINAL_MARKER"]]);

        RequireSuccess(_queries.Update(batch, queryName, newMCode));
        AssertStoredAndEvaluated(queryName, newMCode, ["Marker"], [["NEW_MARKER"]]);
    }

    [Fact]
    public void Update_MultipleSequentialUpdates_EachReplacesCompletely()
    {
        var queryName = "PQ_MultiUpdate_" + Guid.NewGuid().ToString("N")[..8];
        const string version1 = "let V1 = #table({\"Marker\"}, {{\"VERSION_1\"}}) in V1";
        const string version2 = "let V2 = #table({\"Marker\"}, {{\"VERSION_2\"}}) in V2";
        const string version3 = "let V3 = #table({\"Marker\"}, {{\"VERSION_3\"}}) in V3";
        var batch = _fixture.BatchToken;
        RequireSuccess(_queries.Create(batch, queryName, version1,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        AssertStoredAndEvaluated(queryName, version1, ["Marker"], [["VERSION_1"]]);

        RequireSuccess(_queries.Update(batch, queryName, version2));
        AssertStoredAndEvaluated(queryName, version2, ["Marker"], [["VERSION_2"]]);
        RequireSuccess(_queries.Update(batch, queryName, version3));
        AssertStoredAndEvaluated(queryName, version3, ["Marker"], [["VERSION_3"]]);
    }

    [Fact]
    public void Update_InvalidMCodeSyntax_AcceptedByExcel()
    {
        var queryName = $"PQ_Invalid_{Guid.NewGuid():N}"[..20];
        const string validMCode = """
            let
                Source = #table({"A"}, {{1}})
            in
                Source
            """;
        const string invalidMCode = """
            let
                Source = this is not valid M code syntax!!!
            in
                Source
            """;
        var batch = _fixture.BatchToken;
        var guardName = CreateLoadedGuard();
        RequireSuccess(_queries.Create(
            batch,
            queryName,
            validMCode,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        AssertStoredAndEvaluated(queryName, validMCode, ["A"], [[1]]);

        RequireSuccess(_queries.Update(batch, queryName, invalidMCode));
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, invalidMCode,
            PowerQueryLoadMode.ConnectionOnly, null, [], []);
        var before = JsonSerializer.Serialize(RequireSuccess(_queries.List(batch)).Queries);
        var error = Assert.Throws<InvalidOperationException>(() => _queries.Evaluate(batch, invalidMCode));
        Assert.Contains("powerquery.evaluate failed [Expression/PowerQueryCommandException]", error.Message);
        Assert.Equal(before, JsonSerializer.Serialize(RequireSuccess(_queries.List(batch)).Queries));
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, invalidMCode,
            PowerQueryLoadMode.ConnectionOnly, null, [], []);
        AssertInitialTable(guardName);
    }

    [Fact]
    public void Update_ValidMCode_Succeeds()
    {
        var queryName = $"PQ_Valid_{Guid.NewGuid():N}"[..20];
        const string initialMCode =
            "let Source = #table({\"A\"}, {{1}}) in Source";
        const string updatedMCode =
            "let Source = #table({\"A\", \"B\"}, {{1, 2}}) in Source";
        var batch = _fixture.BatchToken;
        RequireSuccess(_queries.Create(
            batch,
            queryName,
            initialMCode,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        AssertStoredAndEvaluated(queryName, initialMCode, ["A"], [[1]]);

        RequireSuccess(_queries.Update(batch, queryName, updatedMCode));
        AssertStoredAndEvaluated(queryName, updatedMCode, ["A", "B"], [[1, 2]]);
    }

    [Fact]
    public void Update_NonExistentQuery_ThrowsWithMeaningfulMessage()
    {
        var queryName = $"PQ_Missing_{Guid.NewGuid():N}"[..20];
        var guardName = CreateLoadedGuard();
        var before = JsonSerializer.Serialize(RequireSuccess(_queries.List(_fixture.BatchToken)).Queries);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _queries.Update(
                _fixture.BatchToken,
                queryName,
                "let Source = 1 in Source"));

        Assert.Contains(queryName, exception.Message);
        Assert.Equal(before, JsonSerializer.Serialize(RequireSuccess(_queries.List(_fixture.BatchToken)).Queries));
        AssertInitialTable(guardName);
    }

    [Fact]
    public void Delete_ExistingQuery_ReturnsSuccess()
    {
        var queryName = "PQ_Delete_" + Guid.NewGuid().ToString("N")[..8];
        var batch = _fixture.BatchToken;
        var guardName = CreateLoadedGuard();
        RequireSuccess(_queries.Create(batch, queryName, TableMCode));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);
        AssertInitialTable(queryName);

        RequireSuccess(_queries.Delete(batch, queryName));
        PowerQueryStateAssertions.AssertRemoved(_fixture, queryName);
        Assert.DoesNotContain(RequireSuccess(_connections.List(batch)).Connections,
            connection => connection.Name == $"Query - {queryName}");
        Assert.All(RequireSuccess(_commands.GetValues(batch, queryName, "A1:C4")).Values,
            row => Assert.All(row, Assert.Null));
        AssertInitialTable(guardName);
        _fixture.ForgetPowerQuery(queryName);
    }

    [Fact]
    public void Create_DuplicateQueryName_ReturnsError()
    {
        const string queryName = "TestQuery";
        var batch = _fixture.BatchToken;
        RequireSuccess(_queries.Create(batch, queryName, TableMCode));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);
        AssertInitialTable(queryName);
        var before = JsonSerializer.Serialize(RequireSuccess(_queries.List(batch)).Queries);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _queries.Create(batch, queryName, "let Source = #table({\"Changed\"}, {{99}}) in Source"));

        Assert.Contains(
            "already exists",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        Assert.Contains(queryName, exception.Message);
        Assert.Equal(before, JsonSerializer.Serialize(RequireSuccess(_queries.List(batch)).Queries));
        AssertInitialTable(queryName);
    }

    private void AssertInitialTable(string name) =>
        PowerQueryStateAssertions.AssertStored(_fixture, name, TableMCode,
            PowerQueryLoadMode.LoadToTable, name, ["Column1", "Column2", "Column3"],
            [["Value1", "Value2", "Value3"], ["A", "B", "C"], ["X", "Y", "Z"]]);

    private string CreateLoadedGuard()
    {
        var name = UniqueCleanupName("LifecycleGuard");
        RequireSuccess(_queries.Create(_fixture.BatchToken, name, TableMCode));
        _fixture.RegisterPowerQueryForCleanup(name);
        _fixture.RegisterSheetForCleanup(name);
        AssertInitialTable(name);
        return name;
    }

    private void AssertStoredAndEvaluated(string name, string code, string[] columns, object[][] rows)
    {
        PowerQueryStateAssertions.AssertStored(_fixture, name, code,
            PowerQueryLoadMode.ConnectionOnly, null, columns, rows);
        var result = RequireSuccess(_queries.Evaluate(_fixture.BatchToken, code));
        Assert.Equal(code, result.MCode);
        Assert.Equal(columns, result.Columns);
        Assert.Equal(columns.Length, result.ColumnCount);
        Assert.Equal(rows.Length, result.RowCount);
        PowerQueryStateAssertions.AssertRows(rows, result.Rows);
        PowerQueryStateAssertions.AssertStored(_fixture, name, code,
            PowerQueryLoadMode.ConnectionOnly, null, columns, rows);
    }
}
