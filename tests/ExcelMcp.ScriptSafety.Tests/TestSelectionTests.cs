using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "PreCommit")]
public sealed class TestSelectionTests
{
    [Theory]
    [InlineData("tests/ExcelMcp.SkillGeneration.Tests/SkillSourceSafetyTests.cs", "SkillGeneration")]
    [InlineData("tests/ExcelMcp.Packaging.Tests/ReleaseMetadataScriptTests.cs", "Packaging")]
    [InlineData("tests/ExcelMcp.ScriptSafety.Tests/ChangedAreaRegressionTests.cs", "ScriptSafety")]
    [InlineData("scripts/Build-AgentSkills.ps1", "SkillGeneration")]
    [InlineData("docs/reference/report-formatting.md", "SkillGeneration")]
    [InlineData("scripts/Build-Plugins.ps1", "Packaging")]
    [InlineData("scripts/Update-ReleaseVersionMetadata.ps1", "Packaging")]
    [InlineData(".github/workflows/publish-mcp-registry.yml", "Packaging")]
    [InlineData("scripts/Resolve-McpRegistryRelease.ps1", "Packaging")]
    [InlineData("scripts/Test-McpRegistryPublication.ps1", "Packaging")]
    [InlineData("scripts/check-workbook-package-access.ps1", "ScriptSafety")]
    [InlineData("tests/Shared/GeneratedAssetsFixture.cs", "SkillGeneration,Packaging")]
    [InlineData("tests/Shared/PackagingScriptTestHelper.cs", "SkillGeneration,Packaging")]
    [InlineData("infrastructure/azure/deploy-appinsights.ps1", "ScriptSafety")]
    [InlineData("videos/excel-mcp-intro/Capture-Evidence.ps1", "ScriptSafety")]
    public async Task ChangedPaths_SelectOwningTestProjects(string path, string expected)
    {
        var result = await RunAsync($$"""
            . (Join-Path $root 'scripts\Get-ValidationPlan.ps1')
            $plan = Get-ValidationPlan -Paths '{{path}}'
            $projects = @(
                if ($plan.SkillTests) { 'SkillGeneration' }
                if ($plan.PackagingTests) { 'Packaging' }
                if ($plan.HookTests) { 'ScriptSafety' }
            )
            if (($projects -join ',') -ne '{{expected}}') { throw "Wrong test projects: $projects" }
            if ($plan.Excel) { throw 'Excel selected for Excel-free changes.' }
            if ($plan.FastProjects.Count -or $plan.ProcessProjects.Count -or $plan.ExcelGroups.Count) {
                throw 'Unrelated runtime tests selected.'
            }
            if ((($plan.ToolingProjects | Sort-Object) -join ',') -ne
                (('{{expected}}'.Split(',') | Sort-Object) -join ',')) {
                throw "Wrong hosted tooling projects: $($plan.ToolingProjects)"
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Theory]
    [InlineData("-Local -SkillTests", "SkillGeneration")]
    [InlineData("-Local -HookTests", "ScriptSafety")]
    [InlineData("-Local -PackagingTests", "Packaging")]
    [InlineData("-Local -SkillTests -HookTests -PackagingTests", "ScriptSafety,SkillGeneration,Packaging")]
    [InlineData("-Local -ChangedPaths @('tests/ExcelMcp.Packaging.Tests/ReleaseMetadataScriptTests.cs')", "Packaging")]
    [InlineData("-Local -ChangedPaths @('scripts/Build-Plugins.ps1')", "Packaging")]
    [InlineData("-Local -ChangedPaths @('scripts/Build-AgentSkills.ps1')", "SkillGeneration")]
    [InlineData("-Local -ChangedPaths @('scripts/Publish-PreparedPlugins.ps1')", "Packaging")]
    [InlineData("-Local -ChangedPaths @('tests/Shared/GeneratedAssetsFixture.cs')", "SkillGeneration,Packaging")]
    [InlineData("-Local -ChangedPaths @('tests/Shared/PackagingScriptTestHelper.cs')", "SkillGeneration,Packaging")]
    [InlineData("", "CLI,ComInterop,Core,McpServer,Service,SkillGeneration,Packaging,ScriptSafety")]
    [InlineData("-Group Tooling", "Packaging,ScriptSafety,SkillGeneration")]
    [InlineData("-Group Tooling -PlanFile $planFile", "Packaging,ScriptSafety,SkillGeneration")]
    public async Task Runner_SelectsActualProjectCommands(string arguments, string expected)
    {
        var result = await RunRunnerAsync(arguments, false);
        Assert.True(result.ExitCode == 0, result.Output);
        Assert.Contains($"selected={expected}", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Runner_SelectedFailureStopsRemainingProjects()
    {
        var result = await RunRunnerAsync("-Local -SkillTests -PackagingTests", true);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("SkillGeneration failed with exit code 23", result.Output, StringComparison.Ordinal);
        Assert.Contains("started=SkillGeneration", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("started=Packaging", result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("'scripts/Build-AgentSkills.ps1'", "SkillGeneration", "SkillSourceSafetyTests")]
    [InlineData("'scripts/Build-Plugins.ps1'", "Packaging", "PluginBootstrap")]
    [InlineData("'scripts/check-workbook-package-access.ps1'", "ScriptSafety", "WorkbookPackageAccessGuardTests")]
    [InlineData("'doc-counts.json'", "Packaging", "FullyQualifiedName~DocumentationCounts")]
    [InlineData("'mcpb/manifest.json'", "Packaging", "McpbPackagingScriptTests")]
    [InlineData("'tests/ExcelMcp.Packaging.Tests/ReleaseMetadataScriptTests.cs'", "Packaging", "ReleaseMetadataScriptTests")]
    [InlineData("'tests/ExcelMcp.ScriptSafety.Tests/ChangedAreaRegressionTests.cs'", "ScriptSafety", "ChangedAreaRegressionTests")]
    [InlineData("'tests/Shared/GeneratedAssetsFixture.cs'", "Packaging,SkillGeneration", "RequiresExcel=false")]
    [InlineData("'tests/Shared/PackagingScriptTestHelper.cs'", "Packaging,SkillGeneration", "RequiresExcel=false")]
    [InlineData("'scripts/Build-AgentSkills.ps1','scripts/check-workbook-package-access.ps1'", "ScriptSafety,SkillGeneration", "SkillSourceSafetyTests")]
    [InlineData("'doc-counts.json','tests/ExcelMcp.ScriptSafety.Tests/ChangedAreaRegressionTests.cs'", "Packaging,ScriptSafety", "FullyQualifiedName~DocumentationCounts")]
    public async Task Runner_HostedToolingSelectsOwningProjectsAndFilters(string paths, string expected, string filter)
    {
        var result = await RunRunnerAsync("-Group Tooling -PlanFile $planFile", false, paths);
        Assert.True(result.ExitCode == 0, result.Output);
        Assert.Contains($"selected={expected}", result.Output, StringComparison.Ordinal);
        Assert.Contains(filter, result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Runner_MixedToolingOwnersUseTheirOwnFilters()
    {
        var result = await RunRunnerAsync("-Group Tooling -PlanFile $planFile", false,
            "'doc-counts.json','tests/ExcelMcp.ScriptSafety.Tests/ChangedAreaRegressionTests.cs'");
        Assert.True(result.ExitCode == 0, result.Output);
        Assert.Contains("Packaging : RequiresExcel=false&RunType!=OnDemand&(FullyQualifiedName~DocumentationCounts)",
            result.Output, StringComparison.Ordinal);
        Assert.Contains("ScriptSafety : RequiresExcel=false&RunType!=OnDemand&(FullyQualifiedName~Sbroenne.ExcelMcp.ScriptSafety.Tests.ChangedAreaRegressionTests.)",
            result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("McpToolSurfaceTests")]
    [InlineData("CalculationGuidanceContractTests")]
    public void Contracts_IncludesExistingMcpContractSuites(string className)
    {
        var selected = FreeTestSelection.Select(TypedValidationPolicyTests.Root,
            new FreeTestOptions { Local = true, Contracts = true });
        Assert.Equal(["Core", "CLI", "McpServer"], selected.Select(item => item.Owner));
        Assert.All(selected, item => Assert.Equal(
            "RequiresExcel=false&RunType!=OnDemand&(Feature=GeneratedContracts)", item.Filter));
        var catalogue = new TestCatalog(TypedValidationPolicyTests.Root);
        var type = Assert.Single(catalogue.ForOwner("McpServer"), type => type.Name == className);
        Assert.True(type.ExcelFree);
        Assert.False(type.Excel);
        Assert.Contains("GeneratedContracts", type.Features);
    }

    private static async Task<(int ExitCode, string Output)> RunRunnerAsync(string arguments, bool fail, string? paths = null)
    {
        var root = TypedValidationPolicyTests.Root;
        var results = Path.Combine(Path.GetTempPath(), $"ExcelMcp.TypedSelection.{Guid.NewGuid():N}");
        var options = new FreeTestOptions
        {
            Local = arguments.Contains("-Local", StringComparison.Ordinal),
            HookTests = arguments.Contains("-HookTests", StringComparison.Ordinal),
            SkillTests = arguments.Contains("-SkillTests", StringComparison.Ordinal),
            PackagingTests = arguments.Contains("-PackagingTests", StringComparison.Ordinal),
            Contracts = arguments.Contains("-Contracts", StringComparison.Ordinal),
            ChangedPaths = Regex.Matches(arguments, "'([^']+)'").Select(match => match.Groups[1].Value).ToArray(),
            Group = arguments.Contains("-Group Tooling", StringComparison.Ordinal) ? "Tooling" : null,
            PlanFile = arguments.Contains("-PlanFile", StringComparison.Ordinal) ? "fixture.json" : null,
            ResultsDirectory = results
        };
        var plan = options.PlanFile is null ? null : new ValidationPolicy(root).Select(
            paths is null ? [] : Regex.Matches(paths, "'([^']+)'").Select(match => match.Groups[1].Value), full: paths is null);
        var runner = new RecordingTestRunner(fail);
        try
        {
            await FreeTestSelection.ExecuteAsync(root, options, runner, plan);
            runner.Output.AppendLine("selected=" + string.Join(',', runner.Owners));
            return (0, runner.Output.ToString());
        }
        catch (InvalidOperationException error) { return (1, runner.Output + error.Message); }
        finally { if (Directory.Exists(results)) { Directory.Delete(results, true); } }
    }

    internal sealed class RecordingTestRunner(bool fail = false) : IProcessRunner
    {
        public StringBuilder Output { get; } = new();
        public List<string> Owners { get; } = [];
        public List<string[]> Commands { get; } = [];

        public Task<ProcessResult> RunAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false)
        {
            var command = arguments.ToArray();
            Commands.Add(command);
            Assert.Equal("dotnet", executable);
            Assert.Equal("test", command[0]);
            Assert.Contains("--disable-build-servers", command);
            Assert.True(deadline > TimeSpan.Zero);
            Assert.False(preserveGitContext);
            var owner = Regex.Match(command[1], @"ExcelMcp\.(\w+)\.Tests\.csproj$").Groups[1].Value;
            Assert.NotEmpty(owner);
            var filter = command[Array.IndexOf(command, "--filter") + 1];
            Assert.Contains("RequiresExcel=false&RunType!=OnDemand", filter, StringComparison.Ordinal);
            var results = command[Array.IndexOf(command, "--results-directory") + 1];
            Assert.NotNull(environment);
            Assert.Equal(Path.Combine(results, $"{owner}-ownership"), environment["EXCELMCP_TEST_OWNERSHIP_DIRECTORY"]);
            Assert.Equal(TypedValidationPolicyTests.Root, environment["EXCELMCP_BUILD_ROOT"]);
            Assert.True(File.Exists(environment["EXCELMCP_BUILD_DLL"]));
            Owners.Add(owner);
            Output.AppendLine("started=" + owner);
            Output.AppendLine(owner + " : " + filter);
            Directory.CreateDirectory(results);
            File.WriteAllText(Path.Combine(results, $"{owner}.trx"),
                """<TestRun><Results><UnitTestResult outcome="Passed"/></Results><ResultSummary outcome="Completed"><Counters total="1" passed="1" executed="1"/></ResultSummary></TestRun>""");
            return Task.FromResult(new ProcessResult(fail ? 23 : 0, "fixture-stdout", "fixture-stderr"));
        }

        public Task<ProcessResult> CheckedAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
            IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false) =>
            throw new InvalidOperationException("Unexpected checked command in a selected test stage.");
    }

    private static async Task<(int ExitCode, string Output)> RunAsync(string body)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !File.Exists(Path.Combine(root.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            root = root.Parent;
        }
        Assert.NotNull(root);
        var sandbox = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"ExcelMcp.Selection.{Guid.NewGuid():N}")).FullName;
        try
        {
            var runner = Path.Combine(sandbox, "test.ps1");
            await File.WriteAllTextAsync(runner, $"""
                $ErrorActionPreference = 'Stop'
                $root = '{root.FullName.Replace("'", "''", StringComparison.Ordinal)}'
                $sandbox = '{sandbox.Replace("'", "''", StringComparison.Ordinal)}'
                {body}
                """);
            var info = new ProcessStartInfo("pwsh")
            {
                WorkingDirectory = sandbox,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            };
            info.Environment["EXCELMCP_BUILD_ROOT"] = root.FullName;
            info.Environment["EXCELMCP_BUILD_DLL"] = typeof(Sbroenne.ExcelMcp.Build.ValidationPolicy).Assembly.Location;
            foreach (var argument in new[] { "-NoProfile", "-File", runner }) { info.ArgumentList.Add(argument); }
            using var process = Process.Start(info)!;
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            try { await process.WaitForExitAsync(deadline.Token); }
            catch (OperationCanceledException)
            {
                process.Kill(true);
                await process.WaitForExitAsync();
                throw new TimeoutException("Test selection exceeded 30 seconds.");
            }
            return (process.ExitCode, await stdout + await stderr);
        }
        finally { Directory.Delete(sandbox, true); }
    }
}
