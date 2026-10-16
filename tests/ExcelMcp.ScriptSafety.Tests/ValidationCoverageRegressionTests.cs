using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ValidationCoverageRegressionTests
{
    private static readonly string Root = TypedValidationPolicyTests.Root;
    private static readonly ValidationPolicy Policy = new(Root);
    private static readonly TestCatalog Catalogue = new(Root);
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PreferredObjectCreationHandling = JsonObjectCreationHandling.Populate
    };
    private const string AzureCases =
        "FullyQualifiedName=Sbroenne.ExcelMcp.ScriptSafety.Tests.AutomationSafetyTests.AzurePreview_NeverMutatesAndRejectsFailedReads|" +
        "FullyQualifiedName=Sbroenne.ExcelMcp.ScriptSafety.Tests.AutomationSafetyTests.AnalyticsOidc_SwitchesToSelectedTenantAndRestoresOriginalSubscription";

    [Theory]
    [InlineData("CLI", "UsageAnalyticsWeightsTests")]
    [InlineData("McpServer", "UsageAnalyticsWeightsTests")]
    [InlineData("McpServer", "McpToolGeneratorTests")]
    [InlineData("McpServer", "StandardToolRegistrationTests")]
    public void ContractSelection_IncludesRegistrationAndAnalyticsSuites(string owner, string name)
    {
        var type = Assert.Single(Catalogue.ForFeatures(owner, ["GeneratedContracts"]), type => type.Name == name);
        Assert.True(type.ExcelFree);
        Assert.False(type.Excel);
        var manual = FreeTestSelection.Select(Root, new FreeTestOptions { Local = true, Contracts = true });
        Assert.Contains("Feature=GeneratedContracts", Assert.Single(manual, item => item.Owner == owner).Filter, StringComparison.Ordinal);
        foreach (var path in new[]
        {
            "src/ExcelMcp.Core/Commands/Range/IRangeCommands.cs",
            "src/ExcelMcp.Generators.Mcp/McpToolGenerator.cs"
        })
        {
            var plan = Policy.Select([path]);
            Assert.Contains($"FullyQualifiedName~{type.FullName}.", plan.FastFilters[owner], StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData("infrastructure/azure/configure-analytics-oidc.ps1")]
    [InlineData("infrastructure/azure/deploy-appinsights.ps1")]
    public async Task AzureChanges_RunSelectedOnDemandCasesThroughLocalAndCiPaths(string path)
    {
        var plan = Policy.Select([path]);
        Assert.False(plan.Excel);
        Assert.Equal(["Tooling"], plan.CiTestGroups);
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(plan));
        Assert.Equal(AzureCases, json.RootElement.GetProperty("ToolingOnDemandFilters").GetProperty("ScriptSafety").GetString());
        var saved = JsonSerializer.Deserialize<ValidationPlan>(json.RootElement.GetRawText(), JsonOptions);
        Assert.NotNull(saved);
        var expectedFilter = $"RequiresExcel=false&((RunType!=OnDemand&({plan.ToolingFilters["ScriptSafety"]}))|(RunType=OnDemand&({AzureCases})))";
        var results = Path.Combine(Path.GetTempPath(), $"ExcelMcp.AzureSelection.{Guid.NewGuid():N}");
        try
        {
            foreach (var options in new[]
            {
                new FreeTestOptions { Local = true, ChangedPaths = [path], ResultsDirectory = results },
                new FreeTestOptions { Group = "Tooling", PlanFile = "saved.json", ResultsDirectory = results }
            })
            {
                var selected = Assert.Single(FreeTestSelection.Select(Root, options, saved));
                Assert.Equal("ScriptSafety", selected.Owner);
                Assert.Equal(expectedFilter, selected.Filter);
                var runner = new TestSelectionTests.RecordingTestRunner();
                await FreeTestSelection.ExecuteAsync(Root, options, runner, saved);
                Assert.Equal(expectedFilter, ReadFilter(Assert.Single(runner.Commands)));
                File.Delete(Path.Combine(results, "ScriptSafety.trx"));
            }
            var validationRunner = new TestSelectionTests.RecordingTestRunner();
            await new ValidationExecution(Root, validationRunner).TestAsync(saved, "Tooling", results);
            Assert.Equal(expectedFilter, ReadFilter(Assert.Single(validationRunner.Commands)));
        }
        finally { if (Directory.Exists(results)) { Directory.Delete(results, true); } }
    }

    [Fact]
    public void OrdinaryAutomationSelection_DoesNotRunOnDemandCases()
    {
        var plan = Policy.Select(["videos/excel-mcp-intro/Capture-Evidence.ps1"]);
        var selected = Assert.Single(FreeTestSelection.Select(Root,
            new FreeTestOptions { Group = "Tooling" }, plan));
        Assert.Equal($"RequiresExcel=false&RunType!=OnDemand&({plan.ToolingFilters["ScriptSafety"]})", selected.Filter);
        Assert.DoesNotContain("RunType=OnDemand", selected.Filter, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, "FullyQualifiedName=Selected.Case")]
    [InlineData(true, "")]
    public void SavedOnDemandSelection_RejectsMissingOwnersAndEmptyFilters(bool ownerSelected, string filter)
    {
        var plan = new ValidationPlan();
        if (ownerSelected) { plan.ToolingFilters["ScriptSafety"] = "FullyQualifiedName~AutomationSafetyTests."; }
        else { plan.ToolingFilters["Packaging"] = "FullyQualifiedName~PluginBuildTests."; }
        plan.ToolingOnDemandFilters["ScriptSafety"] = filter;
        var error = Assert.Throws<InvalidOperationException>(() =>
            FreeTestSelection.Select(Root, new FreeTestOptions { Group = "Tooling" }, plan));
        Assert.Contains("Selected OnDemand cases require", error.Message, StringComparison.Ordinal);
    }

    private static string ReadFilter(string[] arguments) => arguments[Array.IndexOf(arguments, "--filter") + 1];
}
