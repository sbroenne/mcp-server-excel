using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ValidationSelectionTests
{
    [Fact]
    public async Task Paths_SelectExactGroups()
    {
        var run = await RunCasesAsync(
        [
            [".github/dependabot.yml", "", ""],
            ["README.md", "Tooling", ""],
            ["tests/README.md", "", ""],
            ["scripts/Install-CopilotPonytailReview.ps1", "", ""],
            [".github/workflows/copilot-setup-steps.yml", "", ""],
            ["infrastructure/azure/configure-analytics-oidc.ps1", "Tooling", ""],
            ["infrastructure/azure/deploy-appinsights.ps1", "Tooling", ""],
            ["scripts/tests/cli-build-cleanup.tests.ps1", "Tooling", ""],
            ["src/ExcelMcp.Analyzers/SafetyAnalyzer.cs", "Tooling", ""],
            ["doc-counts.json", "Tooling", ""],
            ["scripts/check-doc-counts.ps1", "Tooling", ""],
            ["mcpb/README.md", "Tooling", ""],
            ["vscode-extension/package-lock.json", "", ""],
            ["src/ExcelMcp.CLI/Program.cs", "Fast,Process", "All"],
            ["src/ExcelMcp.Core/Commands/Range/RangeCommands.cs", "Fast", "Range"],
            ["src/ExcelMcp.Core/Commands/PowerQuery/PowerQueryCommands.cs", "Fast", "PowerQuery"],
            ["src/ExcelMcp.Core/Commands/PythonInExcel/PythonInExcelCommands.cs", "Fast", "PythonInExcel"],
            ["src/ExcelMcp.Core/Commands/Calculation/CalculationModeCommands.cs", "Fast", "Calculation"],
            ["src/ExcelMcp.Core/Commands/Table/TableCommands.cs", "Fast", "Table"],
            ["src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.cs", "Fast", "PivotTable"],
            ["src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.cs", "Fast", "DataModel,DataModel dependencies"],
            ["scripts/PluginContent.mjs", "Tooling", ""],
            ["tests/ExcelMcp.Packaging.Tests/PluginPublicationStaging.test.mjs", "Tooling", ""],
            ["tests/ExcelMcp.CLI.Tests/Unit/ActionValidatorTests.cs", "Fast,Process", ""]
        ], """
            $path, $ciGroups, $excelGroups = $case
            $plan = Get-ValidationPlan -Paths $path
            if (($plan.CiTestGroups -join ',') -cne $ciGroups) { throw "Wrong CI groups: $($plan.CiTestGroups)" }
            if (($plan.ExcelGroups -join ',') -cne $excelGroups) { throw "Wrong Excel groups: $($plan.ExcelGroups)" }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task RetiredRunnerRemovals_DoNotSelectProductChecks()
    {
        var run = await RunCasesAsync(
        [
            ["scripts/AzureRunnerHost.ps1"],
            ["scripts/ExcelRunnerHost.ps1"],
            ["scripts/ExcelRunnerPolicy.ps1"],
            ["scripts/Invoke-ExcelRunnerControl.ps1"],
            ["scripts/Invoke-ExcelRunnerMaintenance.ps1"],
            ["scripts/Deploy-ExcelAgentRunner.ps1"],
            ["scripts/Initialize-ExcelAgentDesktop.ps1"],
            ["scripts/Install-ExcelAgentOffice.ps1"],
            ["scripts/Install-ExcelAgentToolchain.ps1"],
            ["scripts/Open-ExcelAgentActivation.ps1"],
            ["scripts/Register-ExcelAgentRunner.ps1"],
            ["scripts/Invoke-CopilotSetupNpm.ps1"]
        ], """
            $plan = Get-ValidationPlan -Paths $case[0]
            if ($plan.Build -or $plan.Excel -or $plan.BuildProjects.Count -or $plan.CiTestGroups.Count) { throw 'Retired operational files selected product checks.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task RetiredRunnerFiles_RequireAnOwningMappingIfRestored()
    {
        var run = await RunAsync("""
            function Test-Path { param([string]$LiteralPath) $true }
            Get-ValidationPlan -Paths 'scripts/AzureRunnerHost.ps1'
            """);
        Assert.NotEqual(0, run.ExitCode);
        Assert.Contains("must not be restored without an owning validation mapping", run.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UnknownRunnerPaths_AreNotTreatedAsRetiredFiles()
    {
        var run = await RunAsync("Get-ValidationPlan -Paths 'scripts/Install-ExcelAgentRunner.ps1'");
        Assert.NotEqual(0, run.ExitCode);
        Assert.Contains("No validation mapping", run.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DocumentationChecks_RunInExactlyOneSelectedGroup()
    {
        var run = await RunCasesAsync(
        [
            ["doc-counts.json", "Tooling"],
            ["scripts/check-doc-counts.ps1", "Tooling"],
            ["src/ExcelMcp.Core/Commands/Range/RangeCommands.Values.cs", ""],
            ["src/ExcelMcp.CLI/Program.cs", ""],
            ["tests/ExcelMcp.CLI.Tests/Unit/ActionValidatorTests.cs", ""],
            ["README.md", ""]
        ], """
            $path, $group = $case
            $plan = Get-ValidationPlan -Paths $path
            if ($plan.DocumentationChecksGroup -cne $group) { throw 'Wrong documentation check group.' }
            if ($group -and $group -notin $plan.CiTestGroups) { throw 'Missing documentation owner.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task CopilotSetup_DoesNotSelectAutomatedValidation()
    {
        var run = await RunCasesAsync(
        [
            ["scripts\\Install-CopilotPonytailReview.ps1"],
            [".github\\workflows\\copilot-setup-steps.yml"]
        ], """
            $plan = Get-ValidationPlan -Paths $case[0]
            if ($plan.Build -or $plan.Excel -or $plan.HookTests -or $plan.CiTestGroups.Count) { throw 'Operational setup selected automated tests.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
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

    [Fact]
    public async Task PackagedSkillInputs_SelectPluginBootstrapChecks()
    {
        var run = await RunCasesAsync(
        [
            ["skills/excel-cli/SKILL.md"],
            ["skills/excel-cli-report-formatting/SKILL.md"],
            ["skills/excel-mcp-report-formatting/SKILL.md"],
            ["scripts/Build-AgentSkills.ps1"],
            ["docs/reference/report-formatting.md"]
        ], """
            $plan = Get-ValidationPlan -Paths $case[0]
            if ($plan.ToolingFilters.Packaging -cne 'Feature=PluginBootstrap') { throw 'Packaged skill inputs must select plugin bootstrap checks.' }
            if ($plan.ToolingFilters.SkillGeneration -cne 'Feature=SkillGeneration') { throw 'Skill generation checks missing.' }
            if (($plan.PackageComponents -join ',') -notmatch 'Plugins') { throw 'Plugin package checks missing.' }
            if ($plan.FastProjects.Count -or $plan.ProcessProjects.Count -or $plan.Excel) { throw 'Unrelated runtime checks selected.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task CaseMatrix_VisitsEveryRowInItsOwnScopeAndReportsTheFailingInput()
    {
        var success = await RunCasesAsync([["first", "1"], ["second'\\path", "2"]], """
            if (Test-Path Variable:matrixLocal) { throw 'Case scope leaked.' }
            $matrixLocal = $true
            Write-Output "visited=$($case[0]):$($case[1])"
            """);
        Assert.True(success.ExitCode == 0, success.Output);
        Assert.Contains("visited=first:1", success.Output, StringComparison.Ordinal);
        Assert.Contains("visited=second'\\path:2", success.Output, StringComparison.Ordinal);

        var failure = await RunCasesAsync([["first"], ["broken"], ["not-reached"]], """
            if ($case[0] -eq 'broken') { throw 'matrix-root-cause' }
            """);
        Assert.NotEqual(0, failure.ExitCode);
        Assert.Contains("Checking case: [\"first\"]", failure.Output, StringComparison.Ordinal);
        Assert.Contains("Checking case: [\"broken\"]", failure.Output, StringComparison.Ordinal);
        Assert.Contains("matrix-root-cause", failure.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("not-reached", failure.Output, StringComparison.Ordinal);
    }

    internal static Task<(int ExitCode, string Output)> RunCasesAsync(string[][] cases, string command)
    {
        Assert.NotEmpty(cases);
        var json = JsonSerializer.Serialize(cases).Replace("'", "''", StringComparison.Ordinal);
        return RunAsync($$"""
            $cases = ConvertFrom-Json -InputObject '{{json}}' -NoEnumerate
            foreach ($case in $cases) {
                Write-Output "Checking case: $(ConvertTo-Json -InputObject $case -Compress)"
                & {
                    {{command}}
                }
            }
            """);
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
        var environment = new Dictionary<string, string>();
        if (environmentVariables is not null)
        {
            foreach (var (name, value) in environmentVariables)
            {
                if (!name.StartsWith("GIT_", StringComparison.OrdinalIgnoreCase)) { environment[name] = value; }
            }
        }
        var result = await new ScriptTestProcess(root.FullName).RunAsync(
            "pwsh", ["-NoProfile", "-Command", $"$ErrorActionPreference='Stop'; . '{script}'; {command}; exit 0"],
            TimeSpan.FromSeconds(120), environment);
        return (result.ExitCode, result.Output + result.Error);
    }
}
