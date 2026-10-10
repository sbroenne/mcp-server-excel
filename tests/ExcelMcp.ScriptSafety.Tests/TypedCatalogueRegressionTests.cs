using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedCatalogueRegressionTests
{
    [Fact]
    public void SharedCoreHelper_SelectsDependentCommandAreasNotTheWholeRuntime()
    {
        var plan = new ValidationPolicy(TypedValidationPolicyTests.Root).Select(["src/ExcelMcp.Core/DataModel/DaxNumericCommaSpacer.cs"]);
        Assert.Contains(plan.ExcelSelections, selection => selection.Filter.Contains("DataModel", StringComparison.Ordinal));
        Assert.DoesNotContain(plan.ExcelSelections, selection => selection.Filter.Contains("Vba", StringComparison.Ordinal));
        Assert.Empty(plan.ProcessProjects);
        Assert.False(plan.FullE2E);
        Assert.False(plan.InfrastructureDiagnostics);
    }

    [Theory]
    [InlineData("Range", "IRangeCommands.cs", "CliNativeRangeAcceptanceTests.")]
    [InlineData("Chart", "IChartConfigCommands.cs", "CliNativeChartAcceptanceTests.")]
    public void ChangedNativeContracts_SelectTheirSpecificExecutableCase(string area, string contract, string required)
    {
        var plan = new ValidationPolicy(TypedValidationPolicyTests.Root).Select([$"src/ExcelMcp.Core/Commands/{area}/{contract}"]);
        Assert.Contains(plan.ExcelSelections, selection => selection.Project == "CLI" && selection.Filter.Contains(required, StringComparison.Ordinal));
        Assert.DoesNotContain(plan.ExcelSelections, selection => selection.Project == "CLI" && selection.Filter.Contains("CliNativePivotAcceptanceTests.", StringComparison.Ordinal));
        Assert.False(plan.FullE2E);
    }

    [Theory]
    [InlineData("src/ExcelMcp.CLI/Commands/ListActionsCommand.cs", "CLI", "ActionValidatorTests.", "CliNativeChartAcceptanceTests.")]
    [InlineData("src/ExcelMcp.McpServer/Telemetry/SensitiveDataRedactor.cs", "McpServer", "TelemetryTests.", "ExcelWorksheetToolTests.")]
    [InlineData("src/ExcelMcp.McpServer/Tools/ExcelFileTool.cs", "McpServer", "ExcelFileToolTests.", "ExcelWorksheetToolTests.")]
    public void AdapterProductionEdits_SelectActualConsumers(string path, string owner, string required, string excluded)
    {
        var plan = new ValidationPolicy(TypedValidationPolicyTests.Root).Select([path]);
        var filters = string.Join('|', plan.FastFilters.Values.Concat(plan.ProcessFilters.Values).Concat(plan.ExcelSelections.Select(selection => selection.Filter)));
        Assert.Contains(required, filters, StringComparison.Ordinal);
        Assert.DoesNotContain(excluded, filters, StringComparison.Ordinal);
        Assert.All(plan.FastProjects.Concat(plan.ProcessProjects), selected => Assert.Equal(owner, selected));
        Assert.All(plan.ExcelSelections, selection => Assert.Equal(owner, selection.Project));
        Assert.False(plan.FullE2E);
    }

    [Fact]
    public void InheritedMethodsAndTraits_BelongToTheConcreteClass()
    {
        WithProject(project =>
        {
            File.WriteAllText(Path.Combine(project, "Cases.cs"), """
                [Trait("RequiresExcel", "false")]
                [Trait("Feature", "Inherited")]
                public abstract class BaseCases { [Fact] public void Check() {} }
                public sealed class ConcreteCases : BaseCases {}
                """);
            var cases = Assert.Single(new TestCatalog(ProjectRoot(project)).ForOwner("Core"));
            Assert.Equal("ConcreteCases", cases.Name);
            Assert.True(cases.ExcelFree);
            Assert.Contains("Inherited", cases.Features);
        });
    }

    [Fact]
    public void CustomFactAttributes_AreDerivedFromTheirActualSource()
    {
        WithProject(project =>
        {
            File.WriteAllText(Path.Combine(project, "Cases.cs"), """
                public class OptionalFactAttribute : FactAttribute {}
                [Trait("RequiresExcel", "false")]
                public class Cases { [OptionalFact] public void Check() {} }
                """);
            Assert.Equal("Cases", Assert.Single(new TestCatalog(ProjectRoot(project)).ForOwner("Core")).Name);
        });
    }

    [Fact]
    public void CollectionFixtureChanges_SelectNamedCollectionConsumers()
    {
        WithProject(project =>
        {
            File.WriteAllText(Path.Combine(project, "Fixture.cs"), "public class Fixture {}");
            File.WriteAllText(Path.Combine(project, "Collection.cs"), """
                [CollectionDefinition("shared state")]
                public class Definition : ICollectionFixture<Fixture> {}
                """);
            File.WriteAllText(Path.Combine(project, "Cases.cs"), """
                [Collection("shared state"), Trait("RequiresExcel", "false")]
                public class Cases { [Fact] public void Check() {} }
                """);
            Assert.Equal("Cases", Assert.Single(new TestCatalog(ProjectRoot(project)).ForFile("tests/ExcelMcp.Core.Tests/Fixture.cs")).Name);
        });
    }

    [Theory]
    [InlineData("public enum SharedInput { First, Second }", "SharedInput value;")]
    [InlineData("public delegate bool SharedInput(string input);", "SharedInput parse;")]
    public void EnumAndDelegateHelpers_SelectConsumers(string helper, string reference)
    {
        WithProject(project =>
        {
            File.WriteAllText(Path.Combine(project, "Helper.cs"), helper);
            File.WriteAllText(Path.Combine(project, "Cases.cs"), $$"""
                [Trait("RequiresExcel", "false")]
                public class Cases { {{reference}} [Fact] public void Check() {} }
                """);
            Assert.Equal("Cases", Assert.Single(new TestCatalog(ProjectRoot(project)).ForFile("tests/ExcelMcp.Core.Tests/Helper.cs")).Name);
        });
    }

    [Fact]
    public void CompileConditions_UseReleaseAndDeclaredProjectSymbols()
    {
        WithProject(project =>
        {
            File.WriteAllText(Path.Combine(project, "ExcelMcp.Core.Tests.csproj"), """
                <Project><PropertyGroup><DefineConstants>$(DefineConstants);LOCAL_CASE</DefineConstants></PropertyGroup></Project>
                """);
            File.WriteAllText(Path.Combine(project, "Cases.cs"), """
                #if DEBUG
                [Trait("RequiresExcel", "false")] public class DebugCases { [Fact] public void Check() {} }
                #endif
                #if !DEBUG && LOCAL_CASE
                [Trait("RequiresExcel", "false")] public class ReleaseCases { [Fact] public void Check() {} }
                #endif
                """);
            Assert.Equal("ReleaseCases", Assert.Single(new TestCatalog(ProjectRoot(project)).ForOwner("Core")).Name);
        });
    }

    [Fact]
    public async Task DeletedHelper_UsesGitSourceToSelectSurvivingConsumers()
    {
        var root = NewProject();
        try
        {
            var project = Path.Combine(root, "tests", "ExcelMcp.Core.Tests");
            var helper = Path.Combine(project, "DifferentFileName.cs");
            File.WriteAllText(helper, "public class SharedFixture {}");
            File.WriteAllText(Path.Combine(project, "Cases.cs"), """
                [Trait("RequiresExcel", "false")]
                public class Cases { SharedFixture fixture; [Fact] public void Check() {} }
                """);
            var git = new ProcessRunner(root);
            await git.CheckedAsync("git", ["init", "--quiet"], TimeSpan.FromSeconds(30));
            await git.CheckedAsync("git", ["add", "."], TimeSpan.FromSeconds(30));
            await git.CheckedAsync("git", ["-c", "user.name=Catalogue fixture", "-c", "user.email=catalogue@example.invalid", "commit", "--quiet", "--no-gpg-sign", "-m", "Catalogue fixture"], TimeSpan.FromSeconds(30));
            File.Delete(helper);
            Assert.Equal("Cases", Assert.Single(new TestCatalog(root).ForFile("tests/ExcelMcp.Core.Tests/DifferentFileName.cs")).Name);
        }
        finally
        {
            foreach (var file in Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)) { File.SetAttributes(file, FileAttributes.Normal); }
            Directory.Delete(root, recursive: true);
        }
    }

    private static void WithProject(Action<string> check)
    {
        var root = NewProject();
        try { check(Path.Combine(root, "tests", "ExcelMcp.Core.Tests")); }
        finally { Directory.Delete(root, recursive: true); }
    }
    private static string ProjectRoot(string project) => Directory.GetParent(project)!.Parent!.FullName;
    private static string NewProject()
    {
        var root = Path.Combine(Path.GetTempPath(), $"ExcelMcp.CatalogueRegression.{Guid.NewGuid():N}");
        var project = Directory.CreateDirectory(Path.Combine(root, "tests", "ExcelMcp.Core.Tests")).FullName;
        File.WriteAllText(Path.Combine(project, "ExcelMcp.Core.Tests.csproj"), "<Project/>");
        return root;
    }
}
