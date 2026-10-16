using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ValidationSelectionTests
{
    private static readonly Sbroenne.ExcelMcp.Build.ValidationPolicy Policy = new(TypedValidationPolicyTests.Root);

    [Theory]
    [InlineData(".github/dependabot.yml", "", "")]
    [InlineData("README.md", "Tooling", "")]
    [InlineData("tests/README.md", "", "")]
    [InlineData("scripts/Install-ExcelAgentToolchain.ps1", "", "")]
    [InlineData("scripts/Invoke-CopilotSetupNpm.ps1", "", "")]
    [InlineData("scripts/Install-CopilotPonytailReview.ps1", "", "")]
    [InlineData(".github/workflows/copilot-setup-steps.yml", "", "")]
    [InlineData("scripts/Register-ExcelAgentRunner.ps1", "", "")]
    [InlineData("infrastructure/azure/update-excel-runner.ps1", "", "")]
    [InlineData(".github/workflows/excel-runner-control.yml", "", "")]
    [InlineData("scripts/ExcelRunnerPolicy.ps1", "", "")]
    [InlineData("scripts/Invoke-ExcelRunnerMaintenance.ps1", "", "")]
    [InlineData("infrastructure/azure/excel-runner.bicep", "", "")]
    [InlineData("infrastructure/azure/configure-analytics-oidc.ps1", "Tooling", "")]
    [InlineData("infrastructure/azure/deploy-appinsights.ps1", "Tooling", "")]
    [InlineData("scripts/tests/excel-runner-build-cleanup.tests.ps1", "Tooling", "")]
    [InlineData("src/ExcelMcp.Analyzers/SafetyAnalyzer.cs", "Tooling", "")]
    [InlineData("doc-counts.json", "Tooling", "")]
    [InlineData("scripts/check-doc-counts.ps1", "Tooling", "")]
    [InlineData("mcpb/README.md", "Tooling", "")]
    [InlineData("vscode-extension/package-lock.json", "", "")]
    [InlineData("src/ExcelMcp.CLI/Program.cs", "Fast,Process", "CLI adapter consumers")]
    [InlineData("src/ExcelMcp.Core/Commands/Range/RangeCommands.cs", "Fast", "Range")]
    [InlineData("src/ExcelMcp.Core/Commands/PowerQuery/PowerQueryCommands.cs", "Fast", "PowerQuery")]
    [InlineData("src/ExcelMcp.Core/Commands/PythonInExcel/PythonInExcelCommands.cs", "Fast", "")]
    [InlineData("src/ExcelMcp.Core/Commands/Calculation/CalculationModeCommands.cs", "Fast", "Calculation")]
    [InlineData("src/ExcelMcp.Core/Commands/Table/TableCommands.cs", "Fast", "Table")]
    [InlineData("src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.cs", "Fast", "PivotTable")]
    [InlineData("src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.cs", "Fast", "DataModel,DataModel dependencies")]
    [InlineData("scripts/PluginContent.mjs", "Tooling", "")]
    [InlineData("tests/ExcelMcp.Packaging.Tests/PluginPublicationHistory.test.mjs", "Tooling", "")]
    [InlineData("tests/ExcelMcp.CLI.Tests/Unit/ActionValidatorTests.cs", "Fast", "")]
    public void Paths_SelectExactGroups(string path, string ciGroups, string excelGroups)
    {
        var plan = Policy.Select([path]);
        Assert.Equal(ciGroups, string.Join(',', plan.CiTestGroups));
        Assert.Equal(excelGroups, string.Join(',', plan.ExcelGroups));
    }

    [Theory]
    [InlineData("doc-counts.json", "Tooling")]
    [InlineData("scripts/check-doc-counts.ps1", "Tooling")]
    [InlineData("src/ExcelMcp.Core/Commands/Range/RangeCommands.Values.cs", "")]
    [InlineData("src/ExcelMcp.CLI/Program.cs", "")]
    [InlineData("tests/ExcelMcp.CLI.Tests/Unit/ActionValidatorTests.cs", "")]
    [InlineData("README.md", "")]
    public void DocumentationChecks_RunInExactlyOneSelectedGroup(string path, string group)
    {
        var plan = Policy.Select([path]);
        Assert.Equal(group, plan.DocumentationChecksGroup);
        if (group.Length > 0) { Assert.Contains(group, plan.CiTestGroups); }
        if (group == "Tooling") { Assert.Equal("FullyQualifiedName~DocumentationCounts", plan.ToolingFilters["Packaging"]); }
    }

    [Theory]
    [InlineData("scripts\\Invoke-CopilotSetupNpm.ps1")]
    [InlineData("scripts\\Install-CopilotPonytailReview.ps1")]
    [InlineData(".github\\workflows\\copilot-setup-steps.yml")]
    public void CopilotSetup_DoesNotSelectAutomatedValidation(string path)
    {
        var plan = Policy.Select([path]);
        Assert.False(plan.Build);
        Assert.False(plan.HookTests);
        Assert.Empty(plan.CiTestGroups);
        Assert.False(plan.Excel);
        Assert.Empty(plan.ExcelGroups);
        Assert.Empty(plan.FastProjects);
        Assert.Empty(plan.ProcessProjects);
    }

    [Fact]
    public async Task CliAcceptanceChanges_RequireOwningAcceptanceValidation()
    {
        var result = await RunAsync("""
            $plan = Get-ValidationPlan -Paths 'tests\ExcelMcp.CLI.Tests\Integration\CliWorkflowAcceptanceTests.cs'
            if (-not $plan.Excel -or -not $plan.Build) { throw 'Owning workbook validation missing.' }
            if ($plan.ExcelSelections.Count -ne 1 -or $plan.ExcelSelections[0].Project -ne 'CLI') {
                throw 'Acceptance selection was broadened beyond its owner.'
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task MixedPaths_UnionGroupsAndNormalizeWindowsPaths()
    {
        var result = await RunAsync("""
            $plan = Get-ValidationPlan -Paths @('src\ExcelMcp.Core\Commands\Range\Deleted.cs', 'src/ExcelMcp.Core/Commands/PowerQuery/Renamed.cs')
            if (($plan.ExcelGroups -join ',') -ne 'PowerQuery,Range') { throw "Wrong union: $($plan.ExcelGroups)" }
            if (($plan.FastProjects -join ',') -ne 'Core,Service') { throw 'Owning Core and Service tests were not selected.' }
            if (-not $plan.ExcelSelections.Count) { throw 'Owning workbook coverage missing.' }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task FullSelection_ContainsEveryHostedTestGroup()
    {
        var result = await RunAsync("""
            $plan = Get-ValidationPlan -Full
            if (($plan.CiTestGroups -join ',') -ne 'Fast,Process,Tooling') { throw 'Full groups missing.' }
            foreach ($project in $plan.ToolingProjects) {
                if (-not $plan.ToolingFilters.$project) { throw 'Full tooling selection has no cases.' }
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task PublicationSelection_DoesNotSelectDaemonOrWorkbookTests()
    {
        var result = await RunAsync("""
            $plan = Get-ValidationPlan -Paths 'scripts/Publish-PreparedPlugins.ps1'
            if ($plan.FastProjects.Count -or $plan.ProcessProjects.Count -or $plan.ExcelGroups.Count) { throw 'Unrelated tests selected.' }
            if ($plan.ToolingFilters.Packaging -notmatch 'PluginPublication') { throw 'Publication regressions missing.' }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    internal static async Task<(int ExitCode, string Output)> RunAsync(
        string command, IReadOnlyDictionary<string, string>? environmentVariables = null)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            root = root.Parent;
        }
        Assert.NotNull(root);
        var script = Path.Combine(root.FullName, "scripts", "Get-ValidationPlan.ps1")
            .Replace("'", "''", StringComparison.Ordinal);
        var environment = new Dictionary<string, string>
        {
            ["EXCELMCP_BUILD_ROOT"] = root.FullName,
            ["EXCELMCP_BUILD_DLL"] = typeof(Sbroenne.ExcelMcp.Build.ValidationPolicy).Assembly.Location
        };
        if (environmentVariables is not null)
        {
            foreach (var (name, value) in environmentVariables)
            {
                if (!name.StartsWith("GIT_", StringComparison.OrdinalIgnoreCase)) { environment[name] = value; }
            }
        }
        var result = await new Sbroenne.ExcelMcp.Build.ProcessRunner(root.FullName).RunAsync(
            "pwsh", ["-NoProfile", "-Command", $"$ErrorActionPreference='Stop'; . '{script}'; {command}"],
            TimeSpan.FromSeconds(120), environment);
        return (result.ExitCode, result.Output + result.Error);
    }
}
