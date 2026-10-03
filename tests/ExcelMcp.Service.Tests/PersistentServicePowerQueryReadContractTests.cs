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
        RequireSuccess(_queries.Create(
            batch,
            queryName,
            mCode,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var list = RequireSuccess(_queries.List(batch));
        var query = Assert.Single(
            list.Queries,
            item => item.Name == queryName);
        var listJson = JsonSerializer.Serialize(list, JsonSerializerOptions.Web);

        Assert.Equal(mCode[..77] + "...", query.FormulaPreview);
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

        var view = RequireSuccess(_queries.View(batch, queryName));

        Assert.Equal(mCode, view.MCode);
        Assert.Equal(mCode.Length, view.CharacterCount);
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, mCode,
            PowerQueryLoadMode.ConnectionOnly, null, ["Value"], [[1]]);
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
                false,
                11),
            new ReadLoadStateScenario(
                $"ReadWorksheet_{suffix}",
                PowerQueryLoadMode.LoadToTable,
                $"ReadWorksheet_{suffix}",
                false,
                false,
                23),
            new ReadLoadStateScenario(
                $"ReadDataModel_{suffix}",
                PowerQueryLoadMode.LoadToDataModel,
                null,
                false,
                true,
                37),
            new ReadLoadStateScenario(
                $"ReadBoth_{suffix}",
                PowerQueryLoadMode.LoadToBoth,
                $"ReadBoth_{suffix}",
                false,
                true,
                49),
        };
        var batch = _fixture.BatchToken;
        foreach (var scenario in scenarios)
        {
            CreateTracked(
                batch,
                scenario.QueryName,
                $"#table({{\"Value\"}}, {{{{{scenario.Seed}}},{{{scenario.Seed + 1}}}}})",
                scenario.LoadMode,
                scenario.TargetSheet);
        }

        var list = RequireSuccess(_queries.List(batch));
        foreach (var scenario in scenarios)
        {
            var query = Assert.Single(
                list.Queries,
                item => item.Name == scenario.QueryName);
            var view = RequireSuccess(_queries.View(batch, scenario.QueryName));
            var loadConfig = RequireSuccess(_queries.GetLoadConfig(batch, scenario.QueryName));

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
            PowerQueryStateAssertions.AssertStored(_fixture, scenario.QueryName,
                $"#table({{\"Value\"}}, {{{{{scenario.Seed}}},{{{scenario.Seed + 1}}}}})",
                scenario.LoadMode, scenario.TargetSheet, ["Value"], [[scenario.Seed], [scenario.Seed + 1]]);
        }
    }

    [Fact]
    public void List_UnexecutedInvalidMQuery_ReturnsCompactMetadata()
    {
        var queryName = $"InvalidRead_{Guid.NewGuid():N}";
        const string invalidMCode =
            "let Source = MissingFunction() in Source";
        var batch = _fixture.BatchToken;
        RequireSuccess(_queries.Create(
            batch,
            queryName,
            invalidMCode,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var result = RequireSuccess(_queries.List(batch));
        var query = Assert.Single(
            result.Queries,
            item => item.Name == queryName);

        Assert.Equal(invalidMCode, query.FormulaPreview);
        Assert.Equal(invalidMCode.Length, query.CharacterCount);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, query.LoadMode);
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, invalidMCode,
            PowerQueryLoadMode.ConnectionOnly, null, [], []);
        var before = JsonSerializer.Serialize(RequireSuccess(_queries.List(batch)).Queries);
        var error = Assert.Throws<InvalidOperationException>(() => _queries.Evaluate(batch, invalidMCode));
        Assert.Contains("MissingFunction", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, JsonSerializer.Serialize(RequireSuccess(_queries.List(batch)).Queries));
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, invalidMCode,
            PowerQueryLoadMode.ConnectionOnly, null, [], []);
    }

    [Fact]
    public void GetLoadConfig_ConnectionOnly_ReturnsConnectionOnlyMode()
    {
        var queryName = "PQ_ConnOnly_" + Guid.NewGuid().ToString("N")[..8];
        var batch = _fixture.BatchToken;
        CreateTracked(
            batch,
            queryName,
            "let Source = #table({\"Val\"}, {{1}}) in Source",
            PowerQueryLoadMode.ConnectionOnly);
        var result = RequireSuccess(_queries.GetLoadConfig(batch, queryName));

        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, result.LoadMode);
        Assert.True(
            string.IsNullOrEmpty(result.TargetSheet),
            "ConnectionOnly should not have a target sheet");
        PowerQueryStateAssertions.AssertStored(_fixture, queryName,
            "let Source = #table({\"Val\"}, {{1}}) in Source",
            PowerQueryLoadMode.ConnectionOnly, null, ["Val"], [[1]]);
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
        CreateTracked(
            batch,
            queryName,
            "let Source = #table({\"Val\"}, {{42}}) in Source",
            PowerQueryLoadMode.LoadToTable,
            sheetName);
        var result = RequireSuccess(_queries.GetLoadConfig(batch, queryName));

        Assert.Equal(PowerQueryLoadMode.LoadToTable, result.LoadMode);
        Assert.Equal(sheetName, result.TargetSheet);
        PowerQueryStateAssertions.AssertStored(_fixture, queryName,
            "let Source = #table({\"Val\"}, {{42}}) in Source",
            PowerQueryLoadMode.LoadToTable, sheetName, ["Val"], [[42]]);
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

        var listResult = RequireSuccess(_queries.List(batch));

        var query = Assert.Single(
            listResult.Queries,
            item => item.Name == queryName);
        Assert.Equal(expectedConnectionOnly, query.IsConnectionOnly);
        PowerQueryStateAssertions.AssertStored(_fixture, queryName,
            "let Source = #table({\"Val\"}, {{1}}) in Source",
            loadMode, targetSheet, ["Val"], [[1]]);
    }

    [Fact]
    public void List_MixedLoadModes_CorrectlyIdentifiesConnectionOnlyQueries()
    {
        var suffix = Guid.NewGuid().ToString("N")[..6];
        var queryConnOnly = "PQ_Mix_ConnOnly_" + suffix;
        var queryTable = "PQ_Mix_Table_" + suffix;
        var queryDataModel = "PQ_Mix_DataModel_" + suffix;
        var queryBoth = "PQ_Mix_Both_" + suffix;
        const string connectionM = "#table({\"A\"}, {{11},{12}})";
        const string tableM = "#table({\"A\"}, {{23},{24}})";
        const string modelM = "#table({\"A\"}, {{37},{38}})";
        const string bothM = "#table({\"A\"}, {{49},{50}})";
        var tableSheet = $"MixTable_{suffix}";
        var bothSheet = $"MixBoth_{suffix}";
        var batch = _fixture.BatchToken;
        CreateTracked(
            batch,
            queryConnOnly,
            connectionM,
            PowerQueryLoadMode.ConnectionOnly);
        CreateTracked(
            batch,
            queryTable,
            tableM,
            PowerQueryLoadMode.LoadToTable,
            tableSheet);
        CreateTracked(
            batch,
            queryDataModel,
            modelM,
            PowerQueryLoadMode.LoadToDataModel);
        CreateTracked(
            batch,
            queryBoth,
            bothM,
            PowerQueryLoadMode.LoadToBoth,
            bothSheet);

        var listResult = RequireSuccess(_queries.List(batch));

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
        PowerQueryStateAssertions.AssertStored(_fixture, queryConnOnly, connectionM,
            PowerQueryLoadMode.ConnectionOnly, null, ["A"], [[11], [12]]);
        PowerQueryStateAssertions.AssertStored(_fixture, queryTable, tableM,
            PowerQueryLoadMode.LoadToTable, tableSheet, ["A"], [[23], [24]]);
        PowerQueryStateAssertions.AssertStored(_fixture, queryDataModel, modelM,
            PowerQueryLoadMode.LoadToDataModel, null, ["A"], [[37], [38]]);
        PowerQueryStateAssertions.AssertStored(_fixture, queryBoth, bothM,
            PowerQueryLoadMode.LoadToBoth, bothSheet, ["A"], [[49], [50]]);
    }

    private void CreateTracked(
        Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch,
        string queryName,
        string mCode,
        PowerQueryLoadMode loadMode,
        string? targetSheet = null)
    {
        RequireSuccess(_queries.Create(batch, queryName, mCode, loadMode, targetSheet));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        if (targetSheet is not null) { _fixture.RegisterSheetForCleanup(targetSheet); }
    }

    private sealed record ReadLoadStateScenario(
        string QueryName,
        PowerQueryLoadMode LoadMode,
        string? TargetSheet,
        bool IsConnectionOnly,
        bool IsLoadedToDataModel,
        int Seed);
}
