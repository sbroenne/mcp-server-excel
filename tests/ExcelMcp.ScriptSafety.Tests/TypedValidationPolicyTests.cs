using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedValidationPolicyTests
{
    internal static readonly string Root = FindRoot();
    private static readonly ValidationPolicy Policy = new(Root);
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PreferredObjectCreationHandling = JsonObjectCreationHandling.Populate
    };

    [Theory]
    [InlineData("src/ExcelMcp.Analyzers/SafetyAnalyzer.cs")]
    [InlineData("src/ExcelMcp.Analyzers/ComAccessAnalyzer.cs")]
    [InlineData("BannedSymbols.txt")]
    [InlineData("BannedSymbols.Mcp.txt")]
    public void CompilerRuleChanges_BuildAllConsumersAndSelectRuleRegressions(string path)
    {
        var plan = Policy.Select([path]);
        Assert.True(plan.FullSolutionBuild);
        Assert.Contains("SafetyAnalyzerTests.", plan.ToolingFilters["ScriptSafety"], StringComparison.Ordinal);
        Assert.False(plan.Excel);
        Assert.False(plan.FullE2E);
        Assert.False(plan.DocumentationCounts);
    }

    [Theory]
    [InlineData("Range", "RangeCommands.Values.cs")]
    [InlineData("PowerQuery", "PowerQueryCommands.Read.cs")]
    public void FeatureEdits_SelectBehaviorWithoutUnchangedAdapters(string area, string file)
    {
        var plan = Policy.Select([$"src/ExcelMcp.Core/Commands/{area}/{file}"]);
        Assert.Equal(["Core"], plan.FastProjects);
        Assert.Empty(plan.ProcessProjects);
        Assert.Empty(plan.ToolingProjects);
        Assert.All(plan.ExcelSelections, selection => Assert.True(selection.Project is "Core" or "Service"));
        Assert.Contains(plan.ExcelSelections, selection => selection.Area == area && selection.Filter.Contains(area, StringComparison.Ordinal));
        Assert.False(plan.FullE2E);
    }

    [Fact]
    public void SharedCliCommandBase_SelectsGeneratedCommandConsumers()
    {
        var plan = Policy.Select(["src/ExcelMcp.CLI/Infrastructure/ServiceCommandBase.cs"]);
        Assert.Equal(["CLI"], plan.FastProjects);
        Assert.Equal(["CLI"], plan.ProcessProjects);
        Assert.Contains(plan.ExcelSelections, selection => selection.Project == "CLI");
        Assert.All(plan.ExcelSelections, selection => Assert.Equal("CLI", selection.Project));
        Assert.False(plan.FullE2E);
    }

    [Fact]
    public void TrayHostEdits_SelectActualDaemonProcessConsumers()
    {
        var plan = Policy.Select(["src/ExcelMcp.CLI/Infrastructure/CliServiceTray.cs"]);
        Assert.Equal(["CLI"], plan.ProcessProjects);
        Assert.Contains("CliDaemonLifecycleTests.", plan.ProcessFilters["CLI"], StringComparison.Ordinal);
        Assert.Contains("CliDaemonTests.", plan.ProcessFilters["CLI"], StringComparison.Ordinal);
        Assert.DoesNotContain("CliDaemonObservationTests.", plan.ProcessFilters["CLI"], StringComparison.Ordinal);
        Assert.Empty(plan.FastProjects);
        Assert.False(plan.Excel || plan.FullE2E);
    }

    [Fact]
    public void BatchAdapterEdit_SelectsBatchEntryPointConsumers()
    {
        var plan = Policy.Select(["src/ExcelMcp.CLI/Commands/BatchCommand.cs"]);

        Assert.Contains(plan.FastFilters.Values.Concat(plan.ProcessFilters.Values),
            filter => filter.Contains("BatchCommandTests.", StringComparison.Ordinal));
        Assert.False(plan.Excel);
        Assert.False(plan.FullE2E);
        Assert.False(plan.FullSolutionBuild);
    }

    [Fact]
    public void TestEdit_SelectsOwningClassAndNotWholeProject()
    {
        var plan = Policy.Select(["tests/ExcelMcp.Core.Tests/Unit/GeneratedActionContractTests.cs"]);
        Assert.Equal(["Core"], plan.FastProjects);
        Assert.Equal("FullyQualifiedName~Sbroenne.ExcelMcp.Core.Tests.Unit.GeneratedActionContractTests.", plan.FastFilters["Core"]);
        Assert.False(plan.Excel);
        Assert.Empty(plan.ProcessProjects);
    }

    [Fact]
    public void TestEdit_DoesNotTreatItsPrivateMocksAsSharedDependencies()
    {
        var plan = Policy.Select(["tests/ExcelMcp.Core.Tests/Unit/RangeHelpersExceptionTests.cs"]);
        Assert.Contains("RangeHelpersExceptionTests.", Assert.Single(plan.FastFilters).Value, StringComparison.Ordinal);
        Assert.DoesNotContain('|', Assert.Single(plan.FastFilters).Value);
        Assert.False(plan.Excel);
    }

    [Theory]
    [InlineData("samples/world-bank-dashboard/prepare_data.py")]
    [InlineData("samples/world-bank-dashboard/world_bank_source.m")]
    [InlineData("samples/world-bank-dashboard/README.md")]
    public void PublishedSampleChanges_SelectDocumentationAndSampleChecksWithoutRuntimeTests(string path)
    {
        var plan = Policy.Select([path]);
        Assert.True(plan.Docs);
        Assert.False(plan.Build);
        Assert.False(plan.Excel);
        Assert.Empty(plan.CiTestGroups);
    }

    [Fact]
    public void EveryExistingCommandArea_HasAnExplicitValidationMapping()
    {
        foreach (var directory in Directory.EnumerateDirectories(Path.Combine(Root, "src", "ExcelMcp.Core", "Commands")))
        {
            var area = Path.GetFileName(directory);
            var plan = Policy.Select([$"src/ExcelMcp.Core/Commands/{area}/Commands.cs"]);
            Assert.True(plan.Build, $"{area} has no build input.");
            Assert.NotEmpty(plan.Reasons);
        }
    }

    [Fact]
    public void MixedChanges_AreAUnionRegardlessOfPathOrder()
    {
        var paths = new[] { "src/ExcelMcp.Core/Commands/Range/RangeCommands.Values.cs", "src/ExcelMcp.Core/Commands/PowerQuery/PowerQueryCommands.Read.cs" };
        var plan = Policy.Select(paths);
        var reversed = Policy.Select(paths.Reverse());
        Assert.NotEmpty(plan.FastFilters);
        Assert.Equal(JsonSerializer.Serialize(plan), JsonSerializer.Serialize(reversed));
    }

    [Theory]
    [InlineData("unknown-build-input.config")]
    [InlineData("src/ExcelMcp.Core/Commands/Unknown/Commands.cs")]
    [InlineData("tests/ExcelMcp.Core.Tests/Unit/Unknown.cs")]
    public void UnknownInputs_FailWithAnOwningAreaDiagnostic(string path)
    {
        var exception = Assert.Throws<InvalidOperationException>(() => Policy.Select([path]));
        Assert.Contains("No validation mapping", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("src/Shared/TelemetryUserIdentity.cs")]
    [InlineData("src/Shared/UsageTelemetry.cs")]
    [InlineData("src/Shared/UsageRecord.cs")]
    public void SharedTelemetrySource_SelectsBothEntryPointIdentityTests(string path)
    {
        var plan = Policy.Select([path]);

        Assert.Equal(["CLI", "McpServer"], plan.FastProjects);
        Assert.Contains("TelemetryUserIdentityTests.", plan.FastFilters["CLI"], StringComparison.Ordinal);
        Assert.Contains("EntryPointTelemetryTests.", plan.FastFilters["CLI"], StringComparison.Ordinal);
        Assert.Contains("TelemetryUserIdentityTests.", plan.FastFilters["McpServer"], StringComparison.Ordinal);
        Assert.Contains("TelemetryTests.", plan.FastFilters["McpServer"], StringComparison.Ordinal);
        Assert.True(plan.Build);
        Assert.False(plan.Excel);
        Assert.False(plan.FullE2E);
    }

    [Theory]
    [InlineData("src/ExcelMcp.CLI/Telemetry/CliTelemetry.cs", "CLI")]
    [InlineData("src/ExcelMcp.McpServer/Telemetry/ExcelMcpTelemetry.cs", "McpServer")]
    public void TelemetryAdapter_SelectsTelemetryWithoutWorkbookFeatures(string path, string owner)
    {
        var plan = Policy.Select([path]);
        Assert.Equal([owner], plan.FastProjects);
        Assert.Contains("TelemetryTests.", plan.FastFilters[owner], StringComparison.Ordinal);
        Assert.True(plan.Build);
        Assert.False(plan.Excel);
        Assert.False(plan.FullE2E);
    }

    [Theory]
    [InlineData("src/ExcelMcp.Analyzers/SafetyAnalyzer.cs")]
    [InlineData("src/ExcelMcp.Analyzers/ComAccessAnalyzer.cs")]
    public void CompilerChanges_SelectTypedRuleRegressions(string path)
    {
        var plan = Policy.Select([path]);
        Assert.Contains("SafetyAnalyzerTests.", plan.ToolingFilters["ScriptSafety"], StringComparison.Ordinal);
        Assert.True(plan.FullSolutionBuild);
        Assert.False(plan.Excel);
    }

    [Theory]
    [InlineData("tools/ExcelMcp.Build/ProcessRunner.cs")]
    [InlineData("tools/ExcelMcp.Build/ExcelMcp.Build.csproj")]
    public void SharedBuildToolInputs_SelectEveryConsumerAndPackagePreparation(string path)
    {
        var plan = Policy.Select([path]);
        Assert.Contains("ScriptSafety", plan.ToolingProjects);
        Assert.Contains("Packaging", plan.ToolingProjects);
        Assert.Contains("SkillGeneration", plan.ToolingProjects);
        Assert.Contains("TypedProcessRunnerTests", plan.ToolingFilters["ScriptSafety"], StringComparison.Ordinal);
        Assert.Contains("TypedPackageOperationsTests", plan.ToolingFilters["Packaging"], StringComparison.Ordinal);
        Assert.False(plan.Excel);
        Assert.False(plan.FullE2E);
    }

    [Fact]
    public void MetadataBundleReadme_SelectsMcpbPackagingInsteadOfGenericDocumentation()
    {
        var plan = Policy.Select(["mcpb/README.md"]);
        Assert.False(plan.Docs);
        Assert.Contains("McpbPackagingScriptTests", plan.ToolingFilters["Packaging"], StringComparison.Ordinal);
    }

    [Fact]
    public void ScopedRuntimeAndSourceGuardChanges_DoNotRequireFullBuildOrDocumentationCounts()
    {
        var range = Policy.Select(["src/ExcelMcp.Core/Commands/Range/RangeCommands.Values.cs"]);
        Assert.True(range.Build);
        Assert.False(range.FullSolutionBuild);
        Assert.False(range.DocumentationCounts);
        Assert.Equal("", range.FullSolutionBuildGroup);

        var guard = Policy.Select(["src/ExcelMcp.Analyzers/ComAccessAnalyzer.cs"]);
        Assert.True(guard.Build);
        Assert.True(guard.FullSolutionBuild);
        Assert.False(guard.DocumentationCounts);
        Assert.Equal("", guard.DocumentationChecksGroup);
        Assert.Equal("Tooling", guard.FullSolutionBuildGroup);
    }

    [Theory]
    [InlineData("../Directory.Build.props")]
    [InlineData("src/../../Directory.Build.props")]
    [InlineData("")]
    public void InvalidPaths_AreRejected(string path) =>
        Assert.Throws<ArgumentException>(() => Policy.Select([path]));

    [Fact]
    public void SavedPlans_RetainFiltersAndSelections()
    {
        var plan = Policy.Select(["src/ExcelMcp.Core/Commands/Range/RangeCommands.Values.cs"]);
        var roundTrip = JsonSerializer.Deserialize<ValidationPlan>(JsonSerializer.Serialize(plan), JsonOptions);
        Assert.NotNull(roundTrip);
        Assert.Equal(plan.FastFilters, roundTrip.FastFilters);
        Assert.Equal(plan.ExcelSelections, roundTrip.ExcelSelections);
        Assert.Equal(plan.BuildProjects, roundTrip.BuildProjects);
        Assert.Equal(plan.Reasons, roundTrip.Reasons);
    }

    [Fact]
    public void Catalogue_ExcludesDiagnosticOnlyClassesFromOrdinaryExecution()
    {
        var catalogue = new TestCatalog(Root);
        var diagnostic = Assert.Single(catalogue.ForOwner("Core"), type => type.Name == "PowerQueryRefreshCpuSpinTests");
        Assert.False(diagnostic.Excel);
        Assert.False(diagnostic.ExcelFree);
    }

    [Fact]
    public void NativeCliWorkflow_DoesNotLaunchAnEmbeddedPowerShellTestRunner()
    {
        var source = File.ReadAllText(Path.Combine(Root, "tests", "ExcelMcp.CLI.Tests", "Integration", "CliWorkflowAcceptanceTests.cs"));
        Assert.DoesNotContain("Test-CliApiCoverage.ps1", source, StringComparison.Ordinal);
    }

    [Fact]
    public void SharedHelperChanges_IncludeTransitiveConsumers()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"ExcelMcp.Catalogue.{Guid.NewGuid():N}");
        Directory.CreateDirectory(Path.Combine(directory, "tests", "ExcelMcp.Core.Tests"));
        try
        {
            var project = Path.Combine(directory, "tests", "ExcelMcp.Core.Tests");
            File.WriteAllText(Path.Combine(project, "ExcelMcp.Core.Tests.csproj"), "<Project/>");
            File.WriteAllText(Path.Combine(project, "Fixture.cs"), "public class Fixture {}");
            File.WriteAllText(Path.Combine(project, "Helper.cs"), "public class Helper { Fixture field; }");
            File.WriteAllText(Path.Combine(project, "Cases.cs"), """
                [Trait("RequiresExcel", "false")]
                public partial class Cases { Helper field; [Fact] public void First() {} }
                """);
            File.WriteAllText(Path.Combine(project, "Cases.More.cs"), "public partial class Cases { [Fact] public void Second() {} }");
            var catalogue = new TestCatalog(directory);
            Assert.Equal("Cases", Assert.Single(catalogue.ForFile("tests/ExcelMcp.Core.Tests/Fixture.cs")).Name);
            Assert.Equal("Cases", Assert.Single(catalogue.ForFile("tests/ExcelMcp.Core.Tests/Cases.More.cs")).Name);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData("bin")]
    [InlineData("obj")]
    public void Catalogue_DoesNotTraverseBuildOutputs(string output)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"ExcelMcp.Catalogue.{Guid.NewGuid():N}");
        var project = Path.Combine(directory, "tests", "ExcelMcp.Core.Tests");
        var artifact = Path.Combine(project, output, "fixture");
        Directory.CreateDirectory(artifact);
        try
        {
            File.WriteAllText(Path.Combine(project, "ExcelMcp.Core.Tests.csproj"), "<Project/>");
            File.WriteAllText(Path.Combine(project, "Cases.cs"),
                """[Trait("RequiresExcel", "false")] public class Cases { [Fact] public void M() {} }""");
            File.WriteAllText(Path.Combine(artifact, "ExcelMcp.Artifact.Tests.csproj"), "Not a source project.");
            var catalogue = new TestCatalog(directory);
            Assert.Equal("Cases", Assert.Single(catalogue.ForOwner("Core")).Name);
            Assert.Empty(catalogue.ForOwner("Artifact"));
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData("scripts/Test-E2E.ps1")]
    [InlineData("scripts/Test-CliWorkflow.ps1")]
    [InlineData("scripts/Test-CliApiCoverage.ps1")]
    [InlineData("scripts/Stop-ExcelCliService.ps1")]
    public void AcceptanceChanges_LeaveRequiredOnlyCasesToTheAcceptanceRunner(string path)
    {
        var plan = Policy.Select([path]);
        Assert.True(plan.FullE2E);
        Assert.True(plan.Excel);
        Assert.True(plan.Build);
        Assert.Equal(
            ["tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj", "tests/ExcelMcp.McpServer.Tests/ExcelMcp.McpServer.Tests.csproj"],
            plan.BuildProjects);
        Assert.False(plan.FullSolutionBuild);
        var catalogue = new TestCatalog(Root);
        foreach (var selection in plan.ExcelSelections)
        {
            var classes = catalogue.ForOwner(selection.Project).ToArray();
            foreach (var filter in selection.Filter.Split('|'))
            {
                var selected = Assert.Single(classes, type => filter == $"FullyQualifiedName~{type.FullName}.");
                Assert.False(selected.RequiredOnly, $"{selected.FullName} would be excluded before the acceptance runner.");
            }
        }
        Assert.DoesNotContain(plan.ExcelSelections, selection => selection.Project == "CLI");
    }

    private static string FindRoot()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln"))) { return directory.FullName; }
        }
        throw new DirectoryNotFoundException("Cannot find the repository root.");
    }
}
