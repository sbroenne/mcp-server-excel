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
        Assert.False(plan.NpmTests);
        Assert.True(plan.Cli && plan.Mcp && plan.Extension);
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
        Assert.False(plan.Packages);
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
        var paths = new[] { "npm-packages/excelcli/package.json", "npm-packages/mcp-server-excel/package.json" };
        var plan = Policy.Select(paths);
        var reversed = Policy.Select(paths.Reverse());
        Assert.True(plan.Cli && plan.Mcp && plan.NpmTests);
        Assert.Equal(JsonSerializer.Serialize(plan), JsonSerializer.Serialize(reversed));
    }

    [Theory]
    [InlineData("src/ExcelMcp.Core/Commands/Range/RangeCommands.Values.cs", false)]
    [InlineData("vscode-extension/src/extension.ts", true)]
    public void ExtensionSourceTests_RunOnlyWhenTheirInputsChanged(string path, bool selected)
    {
        var plan = Policy.Select([path]);
        using var document = JsonDocument.Parse(JsonSerializer.Serialize(plan));
        Assert.Equal(selected, document.RootElement.GetProperty("ExtensionTests").GetBoolean());
        Assert.True(plan.Extension);
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
    [InlineData("scripts/check-com-leaks.ps1")]
    [InlineData("scripts/check-success-flag.ps1")]
    [InlineData("scripts/check-dynamic-casts.ps1")]
    [InlineData("scripts/check-workbook-package-access.ps1")]
    public void SourceGuardChanges_SelectTypedRuleRegressions(string path)
    {
        var plan = Policy.Select([path]);
        Assert.Contains("TypedSourceGuardTests.", plan.ToolingFilters["ScriptSafety"], StringComparison.Ordinal);
        Assert.True(plan.SourceChecks);
        Assert.False(plan.Excel);
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

    [Theory]
    [InlineData("Range", "CliNativeRangeAcceptanceTests")]
    [InlineData("Charts", "CliNativeChartAcceptanceTests")]
    [InlineData("PivotTables", "CliNativePivotAcceptanceTests")]
    public void NativeCliWorkflow_HasIndependentlySelectableOwners(string feature, string name)
    {
        var catalogue = new TestCatalog(Root);
        Assert.Contains(catalogue.ForFeatures("CLI", [feature]), type => type.Name == name && type.Excel);
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
    [InlineData("scripts/Test-E2E.ps1")]
    [InlineData("scripts/Test-CliWorkflow.ps1")]
    [InlineData("scripts/Test-CliApiCoverage.ps1")]
    [InlineData("scripts/Stop-ExcelMcpProcesses.ps1")]
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
