using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

public partial class PowerQueryCommandsTests
{
    [Fact]
    public void List_LongFormula_ReturnsBoundedCompactMetadataWhileViewReturnsFullM()
    {
        var testFile = _fixture.CreateTestFile();
        var queryName = $"CompactRead_{Guid.NewGuid():N}";
        var mCode = BuildLongReadContractMCode();

        using var batch = ExcelSession.BeginBatch(testFile);
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        var list = _powerQueryCommands.List(batch);
        var query = Assert.Single(list.Queries, item => item.Name == queryName);
        var listJson = JsonSerializer.Serialize(list, JsonSerializerOptions.Web);

        Assert.True(list.Success);
#pragma warning disable CS0618
        Assert.Equal(mCode, query.Formula);
#pragma warning restore CS0618
        Assert.InRange(query.FormulaPreview.Length, 1, 80);
        Assert.Equal(mCode.Length, query.CharacterCount);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, query.LoadMode);
        Assert.True(listJson.Length < 1_000, $"Compact list payload was {listJson.Length} characters.");

        using (var document = JsonDocument.Parse(listJson))
        {
            var serializedQuery = Assert.Single(
                document.RootElement.GetProperty("queries").EnumerateArray());
            Assert.False(serializedQuery.TryGetProperty("formula", out _));
        }

        var view = _powerQueryCommands.View(batch, queryName);

        Assert.True(view.Success);
        Assert.Equal(mCode, view.MCode);
        Assert.Equal(mCode.Length, view.CharacterCount);
    }

    [Fact]
    public void ReadActions_AllLoadModes_ReturnTheSameTruthfulLoadState()
    {
        var testFile = _fixture.CreateTestFile();
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
                true)
        };

        using var batch = ExcelSession.BeginBatch(testFile);
        foreach (var scenario in scenarios)
        {
            _powerQueryCommands.Create(
                batch,
                scenario.QueryName,
                "let Source = #table({\"Value\"}, {{1}}) in Source",
                scenario.LoadMode,
                scenario.TargetSheet);
        }

        var list = _powerQueryCommands.List(batch);

        foreach (var scenario in scenarios)
        {
            var query = Assert.Single(list.Queries, item => item.Name == scenario.QueryName);
            var view = _powerQueryCommands.View(batch, scenario.QueryName);
            var loadConfig = _powerQueryCommands.GetLoadConfig(batch, scenario.QueryName);

            Assert.Equal(scenario.IsConnectionOnly, query.IsConnectionOnly);
            Assert.Equal(scenario.IsConnectionOnly, view.IsConnectionOnly);
            Assert.Equal(scenario.LoadMode, query.LoadMode);
            Assert.Equal(scenario.LoadMode, view.LoadMode);
            Assert.Equal(scenario.TargetSheet, query.TargetSheet);
            Assert.Equal(scenario.TargetSheet, view.TargetSheet);
            Assert.Equal(scenario.IsLoadedToDataModel, query.IsLoadedToDataModel);
            Assert.Equal(scenario.IsLoadedToDataModel, view.IsLoadedToDataModel);
            Assert.Equal(!scenario.IsConnectionOnly, view.HasConnection);
            Assert.Equal(scenario.LoadMode, loadConfig.LoadMode);
            Assert.Equal(scenario.TargetSheet, loadConfig.TargetSheet);
            Assert.Equal(scenario.IsLoadedToDataModel, loadConfig.IsLoadedToDataModel);
            Assert.Equal(!scenario.IsConnectionOnly, loadConfig.HasConnection);
        }
    }

    [Fact]
    public void List_UnexecutedInvalidMQuery_ReturnsCompactMetadata()
    {
        var testFile = _fixture.CreateTestFile();
        var queryName = $"InvalidRead_{Guid.NewGuid():N}";
        const string invalidMCode = "let Source = MissingFunction() in Source";

        using var batch = ExcelSession.BeginBatch(testFile);
        _powerQueryCommands.Create(
            batch,
            queryName,
            invalidMCode,
            PowerQueryLoadMode.ConnectionOnly);

        var result = _powerQueryCommands.List(batch);
        var query = Assert.Single(result.Queries, item => item.Name == queryName);

        Assert.True(result.Success);
        Assert.Equal(invalidMCode, query.FormulaPreview);
        Assert.Equal(invalidMCode.Length, query.CharacterCount);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, query.LoadMode);
    }

    private static string BuildLongReadContractMCode()
    {
        var padding = string.Join(
            Environment.NewLine,
            Enumerable.Repeat("// bounded list preview must not serialize this padding", 250));
        return $"let{Environment.NewLine}{padding}{Environment.NewLine}    Source = #table({{\"Value\"}}, {{{{1}}}}){Environment.NewLine}in{Environment.NewLine}    Source";
    }

    private sealed record ReadLoadStateScenario(
        string QueryName,
        PowerQueryLoadMode LoadMode,
        string? TargetSheet,
        bool IsConnectionOnly,
        bool IsLoadedToDataModel);
}
