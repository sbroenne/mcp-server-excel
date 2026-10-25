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
    [Theory]
    [InlineData(PowerQueryLoadMode.LoadToTable)]
    [InlineData(PowerQueryLoadMode.LoadToDataModel)]
    [InlineData(PowerQueryLoadMode.LoadToBoth)]
    public void EmptyLoad_ReportsVerifiedZeroCounts(PowerQueryLoadMode mode)
    {
        var name = $"Empty_{Guid.NewGuid():N}"[..24];
        const string code = "#table(type table [Value = Int64.Type], {})";
        _fixture.RegisterPowerQueryForCleanup(name);
        var created = RequireSuccess(_queries.Create(_fixture.BatchToken, name, code, mode, name));
        Assert.Equal(0, created.RowsLoaded);
        Assert.Equal(mode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth ? 0L : null,
            created.WorksheetRowsLoaded);
        Assert.Equal(mode is PowerQueryLoadMode.LoadToDataModel or PowerQueryLoadMode.LoadToBoth ? 0L : null,
            created.ModelRowsLoaded);
        PowerQueryStateAssertions.AssertStored(_fixture, name, code, mode,
            mode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth ? name : null,
            ["Value"], []);
    }

    [Theory]
    [InlineData(PowerQueryLoadMode.ConnectionOnly, 0)]
    [InlineData(PowerQueryLoadMode.LoadToTable, 3)]
    [InlineData(PowerQueryLoadMode.LoadToDataModel, 3)]
    [InlineData(PowerQueryLoadMode.LoadToBoth, 3)]
    public void LoadMetrics_MatchActualDestinations(PowerQueryLoadMode mode, int expectedRows)
    {
        var name = $"Metrics_{Guid.NewGuid():N}"[..24];
        const string code = "#table(type table [Value = Int64.Type], {{11},{12},{13}})";
        _fixture.RegisterPowerQueryForCleanup(name);
        var created = RequireSuccess(_queries.Create(_fixture.BatchToken, name, code, mode, name));
        var json = JsonSerializer.SerializeToElement(created, created.GetType(), JsonSerializerOptions.Web);
        Assert.True(json.GetProperty("durationMs").GetDouble() > 0);
        Assert.Equal(expectedRows, json.GetProperty("rowsLoaded").GetInt64());
        Assert.Equal(mode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth ? 3L : null,
            created.WorksheetRowsLoaded);
        Assert.Equal(mode is PowerQueryLoadMode.LoadToDataModel or PowerQueryLoadMode.LoadToBoth ? 3L : null,
            created.ModelRowsLoaded);
        PowerQueryStateAssertions.AssertStored(_fixture, name, code, mode,
            mode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth ? name : null,
            ["Value"], [[11], [12], [13]]);
        if (mode == PowerQueryLoadMode.ConnectionOnly) return;

        const string changed = "#table(type table [Value = Int64.Type], {{21},{22}})";
        RequireSuccess(_queries.Update(_fixture.BatchToken, name, changed, refresh: false));
        var refreshed = RequireSuccess(_queries.Refresh(_fixture.BatchToken, name, default));
        var refreshJson = JsonSerializer.SerializeToElement(refreshed, JsonSerializerOptions.Web);
        Assert.True(refreshJson.GetProperty("durationMs").GetDouble() > 0);
        Assert.Equal(2, refreshJson.GetProperty("rowsLoaded").GetInt64());
        PowerQueryStateAssertions.AssertStored(_fixture, name, changed, mode,
            mode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth ? name : null,
            ["Value"], [[21], [22]]);

        var all = RequireSuccess(_queries.RefreshAll(_fixture.BatchToken, refreshPivotTables: false));
        var metrics = Assert.Single(all.QueryMetrics, metric => metric.QueryName == name);
        Assert.Equal(2, metrics.RowsLoaded);
        Assert.True(metrics.DurationMs > 0);

        var loaded = RequireSuccess(_queries.LoadTo(_fixture.BatchToken, name, PowerQueryLoadMode.LoadToTable, name));
        var loadJson = JsonSerializer.SerializeToElement(loaded, loaded.GetType(), JsonSerializerOptions.Web);
        Assert.True(loadJson.GetProperty("durationMs").GetDouble() > 0);
        Assert.Equal(2, loadJson.GetProperty("rowsLoaded").GetInt64());
        PowerQueryStateAssertions.AssertStored(_fixture, name, changed,
            PowerQueryLoadMode.LoadToTable, name, ["Value"], [[21], [22]]);
    }

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
                $"WorksheetTarget_{suffix}",
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
                $"BothTarget_{suffix}",
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
