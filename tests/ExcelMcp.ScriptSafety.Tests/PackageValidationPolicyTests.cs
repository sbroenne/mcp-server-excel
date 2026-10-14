using System.Text.Json;
using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class PackageValidationPolicyTests
{
    [Theory]
    [InlineData("add")]
    [InlineData("remove")]
    [InlineData("version")]
    public async Task EntryPointPackages_SelectTheirOwnersAndDependantsWithoutFullAcceptance(string change)
    {
        using var repo = new Repository();
        repo.Packages(change == "add" ? "" : Package("Adapter", "1.0"));
        repo.Project("CLI", packages: ["Adapter"]);
        repo.Project("McpServer", references: ["CLI"]);
        await repo.CommitAsync();
        repo.Packages(change == "remove" ? "" : Package("adapter", "2.0"));

        var plan = repo.Select();

        Assert.Equal(["CLI", "McpServer"], plan.FastProjects);
        Assert.Equal(["CLI", "McpServer"], plan.ExcelSelections.Select(selection => selection.Project).Distinct().Order());
        Assert.True(plan.Cli && plan.Mcp && plan.Extension && plan.SourceChecks);
        Assert.True(plan.FullSolutionBuild && plan.Build);
        Assert.False(plan.FullE2E);
        Assert.False(plan.InfrastructureDiagnostics);
        Assert.False(plan.SkillTests);
        Assert.DoesNotContain(plan.BuildProjects, path => path.Contains("Core.Tests", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("base-reference")]
    [InlineData("head-reference")]
    [InlineData("removed-project")]
    [InlineData("removed-edge")]
    [InlineData("added-edge")]
    public async Task BothSnapshots_IncludeRemovedConsumersAndTransitiveDependants(string change)
    {
        using var repo = new Repository();
        repo.Packages(Package("Adapter", "1.0"));
        repo.Project("CLI", packages: change == "head-reference" ? [] : ["Adapter"]);
        repo.Project("McpServer", references: change == "added-edge" ? [] : ["CLI"]);
        repo.Project("Diagnostics", references: ["McpServer"]);
        await repo.CommitAsync();
        repo.Packages(Package("Adapter", "2.0"));
        if (change == "base-reference") { repo.Project("CLI"); }
        if (change == "head-reference") { repo.Project("CLI", packages: ["Adapter"]); }
        if (change == "removed-project") { File.Delete(repo.ProjectPath("CLI")); }
        if (change == "removed-edge") { repo.Project("McpServer"); }
        if (change == "added-edge") { repo.Project("McpServer", references: ["CLI"]); }

        var plan = repo.Select();

        Assert.Equal(["CLI", "Diagnostics", "McpServer"], plan.FastProjects);
        Assert.False(plan.FullE2E);
        Assert.True(plan.FullSolutionBuild);
        Assert.DoesNotContain(plan.BuildProjects, path => path.Contains("Core.Tests", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("add")]
    [InlineData("remove")]
    [InlineData("version")]
    [InlineData("comments")]
    public async Task UnreferencedPackages_RequireOnlyASolutionBuild(string change)
    {
        using var repo = new Repository();
        repo.Packages(change == "add" ? "" : Package("Unused", "1.0"));
        await repo.CommitAsync();
        repo.Packages(change == "remove" ? "" : Package("Unused", change == "comments" ? "1.0" : "2.0") + "<!-- explanation -->");

        var plan = repo.Select();

        Assert.True(plan.FullSolutionBuild && plan.Build);
        Assert.False(plan.Excel || plan.FullE2E || plan.SourceChecks || plan.Packages);
        Assert.Empty(plan.CiTestGroups);
        Assert.Empty(plan.BuildProjects);
        Assert.Equal(["csharp"], plan.CodeQlLanguages);
    }

    [Theory]
    [InlineData("Core")]
    [InlineData("ComInterop")]
    [InlineData("Service")]
    public async Task SharedRuntimePackages_RetainFullExcelCoverage(string owner)
    {
        using var repo = new Repository();
        repo.Packages(Package("Runtime", "1.0"));
        repo.Project(owner, packages: ["Runtime"]);
        await repo.CommitAsync();
        repo.Packages(Package("Runtime", "2.0"));

        var plan = repo.Select();

        Assert.True(plan.FullE2E && plan.Excel && plan.InfrastructureDiagnostics);
        Assert.True(plan.FullSolutionBuild && plan.SourceChecks);
        Assert.Equal(["CLI", "ComInterop", "Core", "McpServer", "Service"], plan.FastProjects);
    }

    [Theory]
    [InlineData("<PropertyGroup><ManagePackageVersionsCentrally>false</ManagePackageVersionsCentrally></PropertyGroup>")]
    [InlineData("<Import Project=\"extra.props\" />")]
    [InlineData("<ItemGroup Condition=\"'$(Configuration)' == 'Release'\"><PackageVersion Include=\"Adapter\" Version=\"2.0\" /></ItemGroup>")]
    [InlineData("<ItemGroup><PackageVersion Include=\"Adapter\" Version=\"$(AdapterVersion)\" /></ItemGroup>")]
    [InlineData("<Target Name=\"AdjustPackages\" />")]
    public async Task NonVersionOnlyEdits_RetainFullSelection(string extra)
    {
        using var repo = new Repository();
        repo.Packages(Package("Adapter", "1.0"));
        await repo.CommitAsync();
        repo.Packages(Package("Adapter", "2.0"), extra);

        var plan = repo.Select();

        Assert.True(plan.FullE2E && plan.FullSolutionBuild && plan.InfrastructureDiagnostics);
        Assert.True(plan.SkillTests);
    }

    [Theory]
    [InlineData("Directory.Build.props")]
    [InlineData("Directory.Build.targets")]
    [InlineData("global.json")]
    [InlineData("NuGet.Config")]
    [InlineData("Sbroenne.ExcelMcp.sln")]
    public void OtherSharedBuildInputs_RetainFullSelection(string path)
    {
        using var repo = new Repository();
        var plan = new ValidationPolicy(repo.Root).Select([path]);
        Assert.True(plan.FullE2E && plan.FullSolutionBuild && plan.InfrastructureDiagnostics);
        Assert.True(plan.SkillTests);
    }

    [Fact]
    public async Task UnknownPackageConsumers_StillFailSelection()
    {
        using var repo = new Repository();
        repo.Packages(Package("Unknown", "1.0"));
        repo.Write("tools/Unknown/Unknown.csproj", "<Project><ItemGroup><PackageReference Include=\"Unknown\" /></ItemGroup></Project>");
        await repo.CommitAsync();
        repo.Packages(Package("Unknown", "2.0"));

        var error = Assert.Throws<InvalidOperationException>(() => repo.Select());
        Assert.Contains("No validation mapping", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task SharedTestPackageReferences_SelectTheirTestOwners()
    {
        using var repo = new Repository();
        repo.Packages(Package("TestSdk", "1.0"));
        repo.Write("tests/Directory.Build.props", "<Project><ItemGroup><PackageReference Include=\"TestSdk\" /></ItemGroup></Project>");
        await repo.CommitAsync();
        repo.Packages(Package("TestSdk", "2.0"));

        var plan = repo.Select();

        Assert.Equal(["CLI", "ComInterop", "Core", "Diagnostics", "McpServer", "Service"], plan.FastProjects);
        Assert.Equal(["ScriptSafety", "SkillGeneration"], plan.ToolingProjects);
        Assert.True(plan.Excel);
        Assert.False(plan.FullE2E);
    }

    [Fact]
    public async Task LiteralImports_ResolvePackageAndProjectReferences()
    {
        using var repo = new Repository();
        repo.Packages(Package("Adapter", "1.0"));
        repo.Project("CLI");
        repo.Write("src/ExcelMcp.CLI/ExcelMcp.CLI.csproj", "<Project><Import Project=\"..\\Shared\\References.props\" /></Project>");
        repo.Write("src/Shared/References.props", "<Project><ItemGroup><PackageReference Include=\"Adapter\" /></ItemGroup></Project>");
        await repo.CommitAsync();
        repo.Packages(Package("Adapter", "2.0"));

        var plan = repo.Select();

        Assert.Equal(["CLI"], plan.FastProjects);
        Assert.True(plan.Cli);
        Assert.False(plan.FullE2E || plan.Mcp);
    }

    [Fact]
    public async Task ImportedProjectReferences_AreRelativeToTheConsumingProject()
    {
        using var repo = new Repository();
        repo.Packages(Package("Adapter", "1.0"));
        repo.Project("CLI", packages: ["Adapter"]);
        repo.Write("src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj", "<Project><Import Project=\"..\\..\\build\\References.props\" /></Project>");
        repo.Write("build/References.props", "<Project><ItemGroup><ProjectReference Include=\"..\\ExcelMcp.CLI\\ExcelMcp.CLI.csproj\" /></ItemGroup></Project>");
        await repo.CommitAsync();
        repo.Packages(Package("Adapter", "2.0"));

        var plan = repo.Select();

        Assert.Equal(["CLI", "McpServer"], plan.FastProjects);
        Assert.True(plan.Cli && plan.Mcp && plan.Extension);
        Assert.False(plan.FullE2E);
    }

    [Theory]
    [InlineData("<PackageReference Include=\"$(PackageId)\" />")]
    [InlineData("<ProjectReference Include=\"$(ProjectPath)\" />")]
    [InlineData("<Import Project=\"$(SharedImports)\" />")]
    public async Task UnresolvedProjectInputs_FailExplicitly(string input)
    {
        using var repo = new Repository();
        repo.Packages(Package("Adapter", "1.0"));
        repo.Write("src/ExcelMcp.CLI/ExcelMcp.CLI.csproj", $"<Project><ItemGroup>{input}</ItemGroup></Project>");
        await repo.CommitAsync();
        repo.Packages(Package("Adapter", "2.0"));

        var error = Assert.Throws<InvalidOperationException>(() => repo.Select());
        Assert.Contains("Cannot resolve", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MalformedPackageXml_FailsInsteadOfReturningAPlan()
    {
        using var repo = new Repository();
        repo.Packages(Package("Adapter", "1.0"));
        await repo.CommitAsync();
        repo.Write("Directory.Packages.props", "<Project>");

        Assert.Throws<System.Xml.XmlException>(() => repo.Select());
    }

    [Fact]
    public async Task BaseAndExplicitHead_UseTheirMergeBaseInsteadOfHeadOrWorkingTree()
    {
        using var repo = new Repository();
        repo.Packages(Package("Adapter", "1.0"));
        repo.Project("CLI", packages: ["Adapter"]);
        await repo.CommitAsync();
        await repo.GitAsync("tag", "base");
        repo.Packages(Package("Adapter", "3.0"), "<Import Project=\"unrelated.props\" />");
        await repo.CommitAsync();
        await repo.GitAsync("tag", "advanced-base");
        await repo.GitAsync("checkout", "--quiet", "-b", "package-head", "base");
        repo.Packages(Package("Adapter", "2.0"));
        await repo.CommitAsync();
        repo.Packages(Package("Adapter", "4.0"), "<Target Name=\"Unstaged\" />");

        var options = new CommandOptions(["--base-ref", "advanced-base", "--head-ref", "package-head"]);
        var plan = await Program.ReadPlanAsync(repo.Root, options);
        Assert.Equal(["CLI"], plan.FastProjects);
        Assert.False(plan.FullE2E);
        var direct = await Program.ReadPlanAsync(repo.Root,
            new CommandOptions(["--base-ref", "advanced-base", "--head-ref", "package-head", "--comparison", "Direct"]));
        Assert.True(direct.FullE2E);

        var result = await repo.Runner.CheckedAsync("dotnet",
            [typeof(ValidationPolicy).Assembly.Location, "plan", "--root", repo.Root, "--base-ref", "base", "--head-ref", "package-head"],
            TimeSpan.FromSeconds(30));
        var selected = JsonSerializer.Deserialize<ValidationPlan>(result.Output, Program.JsonOptions)!;
        Assert.Equal(["CLI"], selected.FastProjects);
        Assert.False(selected.FullE2E);
        Assert.True(selected.Cli);
        Assert.Equal(["Directory.Packages.props"], selected.ChangedPaths);
    }

    [Fact]
    public async Task StagedPlan_IgnoresUnstagedPackageAndProjectChanges()
    {
        using var repo = new Repository();
        repo.Packages(Package("Adapter", "1.0"));
        repo.Project("CLI", packages: ["Adapter"]);
        await repo.CommitAsync();
        repo.Packages(Package("Adapter", "2.0"));
        await repo.GitAsync("add", "Directory.Packages.props");
        repo.Packages(Package("Adapter", "3.0"), "<Target Name=\"Unstaged\" />");
        repo.Project("Core", packages: ["Adapter"]);

        var plan = await Program.ReadPlanAsync(repo.Root, new CommandOptions(["--staged"]));

        Assert.Equal(["CLI"], plan.FastProjects);
        Assert.False(plan.FullE2E);
        var free = FreeTestSelection.Select(repo.Root, new FreeTestOptions { Local = true, PlanFile = "saved.json" }, plan);
        Assert.Equal("CLI", Assert.Single(free).Owner);
        Assert.Contains("Cases.CLI.FreeCases.", Assert.Single(free).Filter, StringComparison.Ordinal);
    }

    [Fact]
    public async Task StagedMerge_UsesMergeParentForPackageComparison()
    {
        using var repo = new Repository();
        repo.Packages(Package("Adapter", "1.0") + Package("Runtime", "1.0"));
        repo.Project("CLI", packages: ["Adapter"]);
        repo.Project("Core", packages: ["Runtime"]);
        await repo.CommitAsync();
        repo.Packages(Package("Adapter", "2.0") + Package("Runtime", "1.0"));
        await repo.CommitAsync();
        await repo.GitAsync("tag", "merge-parent");
        repo.Packages(Package("Adapter", "2.0") + Package("Runtime", "2.0"));
        await repo.CommitAsync();
        var mergeParent = await repo.GitAsync("rev-parse", "merge-parent");
        repo.Write(".git/MERGE_HEAD", mergeParent.Output.Trim() + "\n");
        repo.Packages(Package("Adapter", "3.0") + Package("Runtime", "1.0"));
        await repo.GitAsync("add", "Directory.Packages.props");

        var plan = await Program.ReadPlanAsync(repo.Root, new CommandOptions(["--staged"]));

        Assert.Equal(["CLI"], plan.FastProjects);
        Assert.False(plan.FullE2E);
    }

    [Fact]
    public void InvalidComparisonMode_IsRejected() =>
        Assert.Throws<ArgumentException>(() => new CommandOptions(["--comparison", "Guess"]));

    [Fact]
    public async Task PackagePreparation_UsesTheSameCommittedPackageComparison()
    {
        using var repo = new Repository();
        repo.Packages(Package("Adapter", "1.0"));
        repo.Project("CLI", packages: ["Adapter"]);
        await repo.CommitAsync();
        await repo.GitAsync("tag", "base");
        repo.Packages(Package("Adapter", "2.0"));
        await repo.CommitAsync();
        repo.Packages(Package("Adapter", "3.0"), "<Target Name=\"Unstaged\" />");
        var calls = new List<PackageCommand>();
        var commands = new PackageCommands(call =>
        {
            calls.Add(call);
            return repo.Runner.CheckedAsync(call.Executable, call.Arguments, TimeSpan.FromSeconds(30));
        });

        await new PackageExecution(repo.Root, commands).ReleaseAsync(new PackageOptions
        {
            BaseRef = "base",
            Components = ["Plugins"],
            OutputDirectory = Path.Combine(repo.Root, "packages")
        });

        Assert.Equal("git", Assert.Single(calls).Executable);
        Assert.False(Directory.Exists(Path.Combine(repo.Root, "packages")));
    }

    [Theory]
    [InlineData("MergeBase")]
    [InlineData("Direct")]
    public async Task CiAdapter_ForwardsItsComparisonSnapshots(string comparison)
    {
        using var repo = new Repository();
        repo.Packages(Package("Adapter", "1.0"));
        repo.Project("CLI", packages: ["Adapter"]);
        await repo.CommitAsync();
        await repo.GitAsync("tag", "base");
        repo.Packages(Package("Adapter", "2.0"));
        await repo.CommitAsync();
        repo.Packages(Package("Adapter", "3.0"), "<Target Name=\"Unstaged\" />");
        repo.Write("build.ps1", File.ReadAllText(Path.Combine(TypedValidationPolicyTests.Root, "build.ps1")));
        var output = Path.Combine(repo.Root, "plan.json");
        await repo.Runner.CheckedAsync("pwsh",
            ["-NoProfile", "-File", Path.Combine(TypedValidationPolicyTests.Root, "scripts", "Get-CiValidationPlan.ps1"),
             "-BaseRef", "base", "-Comparison", comparison, "-OutputPath", output],
            TimeSpan.FromSeconds(30), new Dictionary<string, string>
            {
                ["EXCELMCP_BUILD_ROOT"] = repo.Root,
                ["EXCELMCP_BUILD_DLL"] = typeof(ValidationPolicy).Assembly.Location,
                ["GITHUB_OUTPUT"] = ""
            });

        var plan = JsonSerializer.Deserialize<ValidationPlan>(await File.ReadAllTextAsync(output), Program.JsonOptions)!;
        Assert.Equal(["CLI"], plan.FastProjects);
        Assert.False(plan.FullE2E);
    }

    [Fact]
    public async Task UnreferencedPackage_SelectsAnEnforcedCiBuildWithoutTests()
    {
        using var repo = new Repository();
        repo.Packages(Package("Unused", "1.0"));
        await repo.CommitAsync();
        await repo.GitAsync("tag", "base");
        repo.Packages(Package("Unused", "2.0"));
        await repo.CommitAsync();
        repo.Write("build.ps1", File.ReadAllText(Path.Combine(TypedValidationPolicyTests.Root, "build.ps1")));
        var output = Path.Combine(repo.Root, "plan.json");
        var githubOutput = Path.Combine(repo.Root, "github-output");
        await repo.Runner.CheckedAsync("pwsh",
            ["-NoProfile", "-File", Path.Combine(TypedValidationPolicyTests.Root, "scripts", "Get-CiValidationPlan.ps1"),
             "-BaseRef", "base", "-OutputPath", output],
            TimeSpan.FromSeconds(30), new Dictionary<string, string>
            {
                ["EXCELMCP_BUILD_ROOT"] = repo.Root,
                ["EXCELMCP_BUILD_DLL"] = typeof(ValidationPolicy).Assembly.Location,
                ["GITHUB_OUTPUT"] = githubOutput
            });

        using var json = JsonDocument.Parse(await File.ReadAllTextAsync(output));
        Assert.True(json.RootElement.TryGetProperty("BuildOnly", out var buildOnly));
        Assert.True(buildOnly.GetBoolean());
        Assert.Contains("build_only=true", await File.ReadAllTextAsync(githubOutput), StringComparison.Ordinal);
        var plan = JsonSerializer.Deserialize<ValidationPlan>(json, Program.JsonOptions)!;
        Assert.Empty(plan.CiTestGroups);
        Assert.Equal(["Sbroenne.ExcelMcp.sln"], new ValidationExecution(repo.Root, repo.Runner).BuildProjects(plan));
    }

    [Theory]
    [InlineData("true", "success", true)]
    [InlineData("false", "skipped", true)]
    [InlineData("true", "skipped", false)]
    [InlineData("true", "failure", false)]
    [InlineData("true", "cancelled", false)]
    [InlineData("false", "success", false)]
    public void CiCompletion_RequiresTheExactBuildOnlyResult(string selected, string result, bool succeeds)
    {
        var error = Record.Exception(() => CiCompletion.Verify(new CiCompletionOptions
        {
            Detection = "success",
            Checks = [
                new("tests", "false", "skipped"), new("packages", "false", "skipped"),
                new("npm", "false", "skipped"), new("build-only", selected, result)
            ]
        }));
        if (succeeds) { Assert.Null(error); }
        else { Assert.IsType<InvalidOperationException>(error); }
    }

    [Fact]
    public void CiBuildOnlyJob_BuildsTheSavedPlanAndReportsToTheRequiredGate()
    {
        var workflow = File.ReadAllText(Path.Combine(TypedValidationPolicyTests.Root, ".github", "workflows", "ci.yml"));
        var job = System.Text.RegularExpressions.Regex.Match(workflow, @"(?ms)^  build-only:\r?\n.*?(?=^  \S|\z)").Value;
        Assert.NotEmpty(job);
        Assert.Contains("if: needs.changes.outputs.build_only == 'true'", job, StringComparison.Ordinal);
        Assert.Contains("./build.ps1 build --plan artifacts/ci-plan.json", job, StringComparison.Ordinal);
        Assert.DoesNotContain("Invoke-ExcelFreeTests", job, StringComparison.Ordinal);
        Assert.Contains("needs: [changes, tests, build-only, packages, npm]", workflow, StringComparison.Ordinal);
        Assert.Contains("BUILD_ONLY: ${{ needs.build-only.result }}", workflow, StringComparison.Ordinal);
        Assert.Contains("-SelectedBuildOnly $env:SELECTED_BUILD_ONLY", workflow, StringComparison.Ordinal);
    }

    private static string Package(string id, string version) => $"<PackageVersion Include=\"{id}\" Version=\"{version}\" />";

    private sealed class Repository : IDisposable
    {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), $"ExcelMcp.PackageSelection.{Guid.NewGuid():N}");
        internal ProcessRunner Runner => new(Root);

        internal Repository()
        {
            Write("Sbroenne.ExcelMcp.sln", "");
            foreach (var owner in new[] { "CLI", "McpServer", "Core", "ComInterop", "Service", "Diagnostics", "ScriptSafety", "SkillGeneration" })
            {
                Write($"tests/ExcelMcp.{owner}.Tests/ExcelMcp.{owner}.Tests.csproj", "<Project />");
                Write($"tests/ExcelMcp.{owner}.Tests/Cases.cs", $$"""
                    using Xunit;
                    namespace Cases.{{owner}};
                    [Trait("RequiresExcel", "false")]
                    public class FreeCases { [Fact] public void Check() {} }
                    [Trait("RequiresExcel", "true")]
                    public class ExcelCases { [Fact] public void Check() {} }
                    """);
            }
        }

        internal void Write(string path, string content)
        {
            var full = Path.Combine(Root, path.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            File.WriteAllText(full, content);
        }

        internal void Packages(string items, string extra = "") =>
            Write("Directory.Packages.props", $"<Project><PropertyGroup><ManagePackageVersionsCentrally>true</ManagePackageVersionsCentrally></PropertyGroup><ItemGroup>{items}</ItemGroup>{extra}</Project>");

        internal string ProjectPath(string owner) => Path.Combine(Root, "src", $"ExcelMcp.{owner}", $"ExcelMcp.{owner}.csproj");

        internal void Project(string owner, string[]? packages = null, string[]? references = null) =>
            Write($"src/ExcelMcp.{owner}/ExcelMcp.{owner}.csproj", $"<Project><ItemGroup>{string.Concat((packages ?? []).Select(id => $"<PackageReference Include=\"{id}\" />"))}{string.Concat((references ?? []).Select(id => $"<ProjectReference Include=\"..\\ExcelMcp.{id}\\ExcelMcp.{id}.csproj\" />"))}</ItemGroup></Project>");

        internal async Task CommitAsync()
        {
            await GitAsync("init", "--quiet");
            await GitAsync("config", "user.name", "Selection Tests");
            await GitAsync("config", "user.email", "selection@example.invalid");
            await GitAsync("add", ".");
            await GitAsync("commit", "--quiet", "-m", "Package selection base");
        }

        internal Task<ProcessResult> GitAsync(params string[] args) => Runner.CheckedAsync("git", args, TimeSpan.FromSeconds(30));
        internal ValidationPlan Select() => new ValidationPolicy(Root).Select(["Directory.Packages.props"]);
        public void Dispose()
        {
            foreach (var file in Directory.EnumerateFiles(Root, "*", SearchOption.AllDirectories))
            {
                File.SetAttributes(file, FileAttributes.Normal);
            }
            Directory.Delete(Root, recursive: true);
        }
    }
}
