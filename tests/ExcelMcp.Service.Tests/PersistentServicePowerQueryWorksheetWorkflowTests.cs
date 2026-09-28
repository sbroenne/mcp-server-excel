using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePowerQueryWorksheetWorkflowTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPowerQueryCommands _queries =
        ServiceCommandProxy.Create<IPowerQueryCommands>(fixture);
    private readonly ISheetCommands _sheets =
        ServiceCommandProxy.Create<ISheetCommands>(fixture);

    [Fact]
    public void Import_QueryReferencingAnotherQuery_LoadsDataSuccessfully()
    {
        var sourceQuery = UniqueName("SourceQuery");
        var derivedQuery = UniqueName("DerivedQuery");
        const string sourceMCode = """
            let
                Source = #table(
                    {"ProductID", "ProductName", "Price"},
                    {
                        {1, "Widget", 10.99},
                        {2, "Gadget", 25.50},
                        {3, "Doohickey", 15.75}
                    }
                )
            in
                Source
            """;
        var derivedMCode = $$"""
            let
                Source = {{sourceQuery}},
                FilteredRows = Table.SelectRows(Source, each [Price] > 15)
            in
                FilteredRows
            """;
        var batch = _fixture.BatchToken;

        CreateTableQuery(sourceQuery, sourceMCode, sourceQuery);
        CreateTableQuery(derivedQuery, derivedMCode, derivedQuery);

        var list = _queries.List(batch);
        Assert.True(list.Success);
        Assert.Equal(2, list.Queries.Count);
        Assert.Contains(list.Queries, query => query.Name == sourceQuery);
        Assert.Contains(list.Queries, query => query.Name == derivedQuery);
        var view = _queries.View(batch, derivedQuery);
        Assert.True(view.Success);
        Assert.Contains(sourceQuery, view.MCode);
        Assert.Contains("Table.SelectRows", view.MCode);
        var sourceRefresh = _queries.Refresh(
            batch,
            sourceQuery,
            TimeSpan.FromMinutes(5));
        Assert.True(
            sourceRefresh.Success,
            $"Source query refresh failed: {sourceRefresh.ErrorMessage}");
        var derivedRefresh = _queries.Refresh(
            batch,
            derivedQuery,
            TimeSpan.FromMinutes(5));
        Assert.True(
            derivedRefresh.Success,
            $"Derived query refresh failed: {derivedRefresh.ErrorMessage}");
    }

    [Fact]
    public void Update_QueryLoadedToSheet_PreservesLoadConfiguration()
    {
        var queryName = UniqueName("LoadedQuery");
        var sheetName = UniqueName("DataSheet");
        var batch = _fixture.BatchToken;
        CreateTableQuery(queryName, InitialTableMCode, sheetName);

        var before = _queries.GetLoadConfig(batch, queryName);
        Assert.True(before.Success, "GetLoadConfig before update failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, before.LoadMode);
        Assert.Equal(sheetName, before.TargetSheet);

        _queries.Update(batch, queryName, UpdatedTableMCode);

        var after = _queries.GetLoadConfig(batch, queryName);
        Assert.True(after.Success, "GetLoadConfig after update failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, after.LoadMode);
        Assert.Equal(sheetName, after.TargetSheet);
        var sheets = _sheets.List(batch);
        Assert.Contains(sheets.Worksheets, sheet => sheet.Name == sheetName);
    }

    [Fact]
    public void UpdateMCodeThenRefresh_QueryLoadedToSheet_PreservesLoadConfiguration()
    {
        var queryName = UniqueName("LoadedQuery");
        var sheetName = UniqueName("DataSheet");
        var batch = _fixture.BatchToken;
        CreateTableQuery(queryName, InitialTableMCode, sheetName);

        var before = _queries.GetLoadConfig(batch, queryName);
        Assert.True(before.Success, "GetLoadConfig before update failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, before.LoadMode);
        Assert.Equal(sheetName, before.TargetSheet);

        _queries.Update(batch, queryName, UpdatedTableMCode);

        var after = _queries.GetLoadConfig(batch, queryName);
        Assert.True(after.Success, "GetLoadConfig after update failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, after.LoadMode);
        Assert.Equal(sheetName, after.TargetSheet);
        Assert.False(
            string.IsNullOrEmpty(after.TargetSheet),
            "Query should have a target sheet (not be connection-only)");
    }

    [Fact]
    public void Update_QueryColumnStructure_UpdatesWorksheetColumns()
    {
        var queryName = UniqueName("ColumnStructure");
        var sheetName = UniqueName("DataSheet");
        var batch = _fixture.BatchToken;
        const string oneColumnMCode = """
            let
                Source = #table(
                    {"Column1"},
                    {{"Value1"}, {"Value2"}}
                )
            in
                Source
            """;
        const string oneColumnUpdatedMCode = """
            let
                Source = #table(
                    {"Column1"},
                    {{"UpdatedValue1"}, {"UpdatedValue2"}, {"UpdatedValue3"}}
                )
            in
                Source
            """;
        const string twoColumnMCode = """
            let
                Source = #table(
                    {"Column1", "Column2"},
                    {{"A", "B"}, {"C", "D"}, {"E", "F"}}
                )
            in
                Source
            """;
        CreateTableQuery(queryName, oneColumnMCode, sheetName);

        var initial = _commands.GetUsedRange(batch, sheetName);
        Assert.True(initial.Success, initial.ErrorMessage);
        Assert.Equal(1, initial.ColumnCount);
        _queries.Update(batch, queryName, oneColumnUpdatedMCode);
        var afterFirstUpdate = _commands.GetUsedRange(batch, sheetName);
        Assert.True(afterFirstUpdate.Success, afterFirstUpdate.ErrorMessage);
        Assert.Equal(1, afterFirstUpdate.ColumnCount);

        _queries.Update(batch, queryName, twoColumnMCode);

        var afterSecondUpdate = _commands.GetUsedRange(batch, sheetName);
        Assert.True(afterSecondUpdate.Success, afterSecondUpdate.ErrorMessage);
        var values = _commands.GetValues(
            batch,
            sheetName,
            afterSecondUpdate.RangeAddress);
        Assert.True(values.Success, values.ErrorMessage);
        var headers = values.Values.FirstOrDefault();
        var columnNames = headers is null
            ? "No headers found"
            : string.Join(", ", headers.Select(value => value?.ToString() ?? "null"));
        Assert.True(
            afterSecondUpdate.ColumnCount == 2,
            $"Expected 2 columns but got {afterSecondUpdate.ColumnCount}. " +
            $"Actual columns: [{columnNames}]");
        Assert.True(
            values.ColumnCount == 2,
            $"Expected 2 columns in values but got {values.ColumnCount}. " +
            $"Columns: [{columnNames}]");
    }

    [Fact]
    public void Update_QueryColumnStructureWithDeleteRecreate_NoAccumulation()
    {
        var queryName = UniqueName("AccumulationBug");
        var sheetName = UniqueName("TestSheet");
        var batch = _fixture.BatchToken;
        const string oneColumnMCode = """
            let
                Source = #table({"Column1"}, {{"A"}, {"B"}})
            in
                Source
            """;
        const string twoColumnMCode = """
            let
                Source = #table(
                    {"Column1", "Column2"},
                    {{"X", "Y"}, {"Z", "W"}}
                )
            in
                Source
            """;
        CreateTableQuery(queryName, oneColumnMCode, sheetName);
        var initial = _commands.GetUsedRange(batch, sheetName);
        Assert.True(initial.Success);
        Assert.Equal(1, initial.ColumnCount);

        _queries.Update(batch, queryName, twoColumnMCode);
        _queries.Delete(batch, queryName);
        _fixture.ForgetPowerQuery(queryName);
        _queries.Create(
            batch,
            queryName,
            twoColumnMCode,
            PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _queries.LoadTo(
            batch,
            queryName,
            PowerQueryLoadMode.LoadToTable,
            sheetName,
            "A1");
        var refresh = _queries.Refresh(
            batch,
            queryName,
            TimeSpan.FromMinutes(5));
        Assert.True(refresh.Success, $"Refresh failed: {refresh.ErrorMessage}");

        var usedRange = _commands.GetUsedRange(batch, sheetName);
        Assert.True(usedRange.Success);
        var values = _commands.GetValues(
            batch,
            sheetName,
            usedRange.RangeAddress);
        Assert.True(values.Success);
        var headers = values.Values.FirstOrDefault();
        var columnNames = headers is null
            ? "No headers found"
            : string.Join(", ", headers.Select(value => value?.ToString() ?? "null"));
        Assert.True(
            usedRange.ColumnCount == 2,
            $"COLUMN ACCUMULATION DETECTED! Expected 2 columns but got " +
            $"{usedRange.ColumnCount}. Actual columns: [{columnNames}].");
    }

    private void CreateTableQuery(
        string queryName,
        string mCode,
        string sheetName)
    {
        _queries.Create(
            _fixture.BatchToken,
            queryName,
            mCode,
            PowerQueryLoadMode.LoadToTable,
            sheetName);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(sheetName);
    }

    private static string UniqueName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..Math.Min(prefix.Length + 9, 31)];

    private const string InitialTableMCode = """
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

    private const string UpdatedTableMCode = """
        let
            UpdatedSource = #table(
                {"NewCol1", "NewCol2"},
                {{"Updated1", "Updated2"}, {"Data1", "Data2"}}
            )
        in
            UpdatedSource
        """;
}
