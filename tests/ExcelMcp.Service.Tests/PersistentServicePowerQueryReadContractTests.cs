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
public sealed class PersistentServicePowerQueryReadContractTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPowerQueryCommands _queries =
        ServiceCommandProxy.Create<IPowerQueryCommands>(fixture);

    [Fact]
    public void List_LongFormula_ReturnsBoundedCompactMetadataWhileViewReturnsFullM()
    {
        var queryName = $"CompactRead_{Guid.NewGuid():N}";
        var mCode = BuildLongReadContractMCode();
        var batch = _fixture.BatchToken;
        _queries.Create(
            batch,
            queryName,
            mCode,
            PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var list = _queries.List(batch);
        var query = Assert.Single(
            list.Queries,
            item => item.Name == queryName);
        var listJson = JsonSerializer.Serialize(list, JsonSerializerOptions.Web);

        Assert.True(list.Success);
        Assert.InRange(query.FormulaPreview.Length, 1, 80);
        Assert.Equal(mCode.Length, query.CharacterCount);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, query.LoadMode);
        Assert.True(
            listJson.Length < 1_000,
            $"Compact list payload was {listJson.Length} characters.");

        using (var document = JsonDocument.Parse(listJson))
        {
            var serializedQuery = Assert.Single(
                document.RootElement.GetProperty("queries").EnumerateArray());
            Assert.False(serializedQuery.TryGetProperty("formula", out _));
        }

        var view = _queries.View(batch, queryName);

        Assert.True(view.Success);
        Assert.Equal(mCode, view.MCode);
        Assert.Equal(mCode.Length, view.CharacterCount);
    }

    [Fact]
    public void ReadActions_AllLoadModes_ReturnTheSameTruthfulLoadState()
    {
        var suffix = Guid.NewGuid().ToString("N")[..8];
        var scenarios = new[]
        {
            new ReadLoadStateScenario(
                $"ReadConnectionOnly_{suffix}",
                PowerQueryLoadMode.ConnectionOnly,
                null,
                true,
                false),
            new ReadLoadStateScenario(
                $"ReadWorksheet_{suffix}",
                PowerQueryLoadMode.LoadToTable,
                $"ReadWorksheet_{suffix}",
                false,
                false),
            new ReadLoadStateScenario(
                $"ReadDataModel_{suffix}",
                PowerQueryLoadMode.LoadToDataModel,
                null,
                false,
                true),
            new ReadLoadStateScenario(
                $"ReadBoth_{suffix}",
                PowerQueryLoadMode.LoadToBoth,
                $"ReadBoth_{suffix}",
                false,
                true),
        };
        var batch = _fixture.BatchToken;
        foreach (var scenario in scenarios)
        {
            _queries.Create(
                batch,
                scenario.QueryName,
                "let Source = #table({\"Value\"}, {{1}}) in Source",
                scenario.LoadMode,
                scenario.TargetSheet);
            _fixture.RegisterPowerQueryForCleanup(scenario.QueryName);
        }

        var list = _queries.List(batch);
        foreach (var scenario in scenarios)
        {
            var query = Assert.Single(
                list.Queries,
                item => item.Name == scenario.QueryName);
            var view = _queries.View(batch, scenario.QueryName);
            var loadConfig = _queries.GetLoadConfig(batch, scenario.QueryName);

            Assert.Equal(scenario.IsConnectionOnly, query.IsConnectionOnly);
            Assert.Equal(scenario.IsConnectionOnly, view.IsConnectionOnly);
            Assert.Equal(scenario.LoadMode, query.LoadMode);
            Assert.Equal(scenario.LoadMode, view.LoadMode);
            Assert.Equal(scenario.TargetSheet, query.TargetSheet);
            Assert.Equal(scenario.TargetSheet, view.TargetSheet);
            Assert.Equal(
                scenario.IsLoadedToDataModel,
                query.IsLoadedToDataModel);
            Assert.Equal(
                scenario.IsLoadedToDataModel,
                view.IsLoadedToDataModel);
            Assert.Equal(!scenario.IsConnectionOnly, view.HasConnection);
            Assert.Equal(scenario.LoadMode, loadConfig.LoadMode);
            Assert.Equal(scenario.TargetSheet, loadConfig.TargetSheet);
            Assert.Equal(
                scenario.IsLoadedToDataModel,
                loadConfig.IsLoadedToDataModel);
            Assert.Equal(!scenario.IsConnectionOnly, loadConfig.HasConnection);
        }
    }

    [Fact]
    public void List_UnexecutedInvalidMQuery_ReturnsCompactMetadata()
    {
        var queryName = $"InvalidRead_{Guid.NewGuid():N}";
        const string invalidMCode =
            "let Source = MissingFunction() in Source";
        var batch = _fixture.BatchToken;
        _queries.Create(
            batch,
            queryName,
            invalidMCode,
            PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var result = _queries.List(batch);
        var query = Assert.Single(
            result.Queries,
            item => item.Name == queryName);

        Assert.True(result.Success);
        Assert.Equal(invalidMCode, query.FormulaPreview);
        Assert.Equal(invalidMCode.Length, query.CharacterCount);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, query.LoadMode);
    }

    [Fact]
    public void GetLoadConfig_ConnectionOnly_ReturnsConnectionOnlyMode()
    {
        var queryName = "PQ_ConnOnly_" + Guid.NewGuid().ToString("N")[..8];
        var batch = _fixture.BatchToken;
        _queries.Create(
            batch,
            queryName,
            "let Source = #table({\"Val\"}, {{1}}) in Source",
            PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var result = _queries.GetLoadConfig(batch, queryName);

        Assert.True(result.Success, $"GetLoadConfig failed: {result.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, result.LoadMode);
        Assert.True(
            string.IsNullOrEmpty(result.TargetSheet),
            "ConnectionOnly should not have a target sheet");
    }

    private static string BuildLongReadContractMCode()
    {
        var padding = string.Join(
            Environment.NewLine,
            Enumerable.Repeat(
                "// bounded list preview must not serialize this padding",
                250));
        return
            $"let{Environment.NewLine}{padding}{Environment.NewLine}" +
            $"    Source = #table({{\"Value\"}}, {{{{1}}}})" +
            $"{Environment.NewLine}in{Environment.NewLine}    Source";
    }

    [Fact]
    public void GetLoadConfig_LoadToTable_ReturnsLoadToTableMode()
    {
        var queryName = "PQ_Table_" + Guid.NewGuid().ToString("N")[..8];
        const string sheetName = "TableSheet";
        var batch = _fixture.BatchToken;
        _queries.Create(
            batch,
            queryName,
            "let Source = #table({\"Val\"}, {{42}}) in Source",
            PowerQueryLoadMode.LoadToTable,
            sheetName);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var result = _queries.GetLoadConfig(batch, queryName);

        Assert.True(result.Success, $"GetLoadConfig failed: {result.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, result.LoadMode);
        Assert.Equal(sheetName, result.TargetSheet);
    }

    [Fact]
    public void List_DataModelOnly_NotReportedAsConnectionOnly() =>
        AssertConnectionOnlyState(PowerQueryLoadMode.LoadToDataModel, false);

    [Fact]
    public void List_ConnectionOnly_ReportsAsConnectionOnly() =>
        AssertConnectionOnlyState(PowerQueryLoadMode.ConnectionOnly, true);

    [Fact]
    public void List_LoadToTable_NotReportedAsConnectionOnly() =>
        AssertConnectionOnlyState(PowerQueryLoadMode.LoadToTable, false);

    [Fact]
    public void List_LoadToBoth_NotReportedAsConnectionOnly() =>
        AssertConnectionOnlyState(PowerQueryLoadMode.LoadToBoth, false);

    private void AssertConnectionOnlyState(
        PowerQueryLoadMode loadMode,
        bool expectedConnectionOnly)
    {
        var queryName = $"PQ_List_{loadMode}_{Guid.NewGuid():N}"[..30];
        var targetSheet = loadMode is PowerQueryLoadMode.LoadToTable
            or PowerQueryLoadMode.LoadToBoth
            ? $"Target_{Guid.NewGuid():N}"[..30]
            : null;
        var batch = _fixture.BatchToken;
        CreateTracked(
            batch,
            queryName,
            "let Source = #table({\"Val\"}, {{1}}) in Source",
            loadMode,
            targetSheet);

        var listResult = _queries.List(batch);

        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        var query = Assert.Single(
            listResult.Queries,
            item => item.Name == queryName);
        Assert.Equal(expectedConnectionOnly, query.IsConnectionOnly);
    }

    [Fact]
    public void List_MixedLoadModes_CorrectlyIdentifiesConnectionOnlyQueries()
    {
        var suffix = Guid.NewGuid().ToString("N")[..6];
        var queryConnOnly = "PQ_Mix_ConnOnly_" + suffix;
        var queryTable = "PQ_Mix_Table_" + suffix;
        var queryDataModel = "PQ_Mix_DataModel_" + suffix;
        var queryBoth = "PQ_Mix_Both_" + suffix;
        const string mCode = "let Source = #table({\"A\"}, {{1}}) in Source";
        var batch = _fixture.BatchToken;
        CreateTracked(
            batch,
            queryConnOnly,
            mCode,
            PowerQueryLoadMode.ConnectionOnly);
        CreateTracked(
            batch,
            queryTable,
            mCode,
            PowerQueryLoadMode.LoadToTable,
            "Sheet1");
        CreateTracked(
            batch,
            queryDataModel,
            mCode,
            PowerQueryLoadMode.LoadToDataModel);
        CreateTracked(
            batch,
            queryBoth,
            mCode,
            PowerQueryLoadMode.LoadToBoth,
            "Sheet2");

        var listResult = _queries.List(batch);

        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        var connectionOnly = Assert.Single(
            listResult.Queries,
            query => query.Name == queryConnOnly);
        var table = Assert.Single(
            listResult.Queries,
            query => query.Name == queryTable);
        var dataModel = Assert.Single(
            listResult.Queries,
            query => query.Name == queryDataModel);
        var both = Assert.Single(
            listResult.Queries,
            query => query.Name == queryBoth);
        Assert.True(connectionOnly.IsConnectionOnly);
        Assert.False(table.IsConnectionOnly);
        Assert.False(dataModel.IsConnectionOnly);
        Assert.False(both.IsConnectionOnly);
    }

    private void CreateTracked(
        Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch,
        string queryName,
        string mCode,
        PowerQueryLoadMode loadMode,
        string? targetSheet = null)
    {
        _queries.Create(batch, queryName, mCode, loadMode, targetSheet);
        _fixture.RegisterPowerQueryForCleanup(queryName);
    }

    private sealed record ReadLoadStateScenario(
        string QueryName,
        PowerQueryLoadMode LoadMode,
        string? TargetSheet,
        bool IsConnectionOnly,
        bool IsLoadedToDataModel);
}
