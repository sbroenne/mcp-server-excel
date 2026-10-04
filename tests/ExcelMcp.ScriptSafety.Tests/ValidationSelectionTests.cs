using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ValidationSelectionTests
{
    [Theory]
    [InlineData(".github/dependabot.yml", "", "", false)]
    [InlineData("README.md", "Tooling", "", true)]
    [InlineData("tests/README.md", "", "", false)]
    [InlineData("scripts/Install-ExcelAgentToolchain.ps1", "Tooling", "", false)]
    [InlineData("scripts/Invoke-CopilotSetupNpm.ps1", "Tooling", "", false)]
    [InlineData(".github/workflows/copilot-setup-steps.yml", "Tooling", "", false)]
    [InlineData("scripts/tests/excel-runner-maintenance.tests.ps1", "Tooling", "", false)]
    [InlineData("infrastructure/azure/update-excel-runner.ps1", "Tooling", "", false)]
    [InlineData("doc-counts.json", "Tooling", "", false)]
    [InlineData("scripts/check-doc-counts.ps1", "Tooling", "", false)]
    [InlineData("vscode-extension/package-lock.json", "", "", true)]
    [InlineData("src/ExcelMcp.CLI/Program.cs", "Fast,Process", "Acceptance,Lifecycle", true)]
    [InlineData("src/ExcelMcp.Core/Commands/Range/RangeCommands.cs", "Fast,Process", "Acceptance,Editing", true)]
    [InlineData("src/ExcelMcp.Core/Commands/PowerQuery/PowerQueryCommands.cs", "Fast,Process", "Acceptance,Data", true)]
    [InlineData("src/ExcelMcp.Core/Commands/PythonInExcel/PythonInExcelCommands.cs", "Fast,Process", "Acceptance,Data", true)]
    [InlineData("src/ExcelMcp.Core/Commands/Calculation/CalculationModeCommands.cs", "Fast,Process", "Acceptance,Reporting", true)]
    [InlineData("src/ExcelMcp.Core/Commands/Table/TableCommands.cs", "Fast,Process", "Acceptance,Data,Editing", true)]
    [InlineData("src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.cs", "Fast,Process", "Acceptance,Data,Reporting", true)]
    [InlineData("src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.cs", "Fast,Process", "Acceptance,Data,Editing,Reporting", true)]
    [InlineData("scripts/PluginContent.mjs", "Tooling", "", true)]
    [InlineData("tests/ExcelMcp.Packaging.Tests/PluginPublicationHistory.test.mjs", "Tooling", "", false)]
    [InlineData("tests/ExcelMcp.CLI.Tests/Unit/ActionValidatorTests.cs", "Fast,Process", "", false)]
    [InlineData("tests/Shared/TestRunExcelLifetime.cs", "Fast,Process,Tooling", "Acceptance,Data,Desktop,Editing,Infrastructure,Lifecycle,Reporting,VBA", false)]
    [InlineData("unknown-build-input.config", "Fast,Process,Tooling", "Acceptance,Data,Desktop,Editing,Infrastructure,Lifecycle,Reporting,VBA", true)]
    public async Task Paths_SelectExactGroups(string path, string ciGroups, string excelGroups, bool packages)
    {
        var result = await RunAsync($$"""
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if (($plan.CiTestGroups -join ',') -ne '{{ciGroups}}') { throw "Wrong CI groups: $($plan.CiTestGroups)" }
            if (($plan.ExcelGroups -join ',') -ne '{{excelGroups}}') { throw "Wrong Excel groups: $($plan.ExcelGroups)" }
            if ($plan.Packages -ne ${{(packages ? "true" : "false")}}) { throw 'Wrong package selection.' }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Theory]
    [InlineData("doc-counts.json", "Tooling")]
    [InlineData("scripts/check-doc-counts.ps1", "Tooling")]
    [InlineData("src/ExcelMcp.CLI/Program.cs", "Fast")]
    [InlineData("tests/ExcelMcp.CLI.Tests/Unit/ActionValidatorTests.cs", "")]
    [InlineData("README.md", "")]
    public async Task SourceChecks_RunInExactlyOneSelectedGroup(string path, string group)
    {
        var result = await RunAsync($$"""
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if ($plan.SourceChecksGroup -cne '{{group}}') { throw "Wrong source-check group: $($plan.SourceChecksGroup)" }
            if ($plan.SourceChecksGroup -and $plan.SourceChecksGroup -notin $plan.CiTestGroups) { throw 'Source checks have no selected job.' }
            if ('{{group}}' -eq 'Tooling' -and $plan.ToolingFilters.Packaging -cne 'FullyQualifiedName~DocumentationCounts') {
                throw 'Documentation-count regressions were not selected.'
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Theory]
    [InlineData("scripts\\Invoke-CopilotSetupNpm.ps1")]
    [InlineData(".github\\workflows\\copilot-setup-steps.yml")]
    public async Task CopilotNpmSetup_SelectsSafetyWithoutExcel(string path)
    {
        var result = await RunAsync($$"""
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if (-not $plan.Build -or -not $plan.HookTests) { throw 'Script safety validation missing.' }
            if ($plan.Excel -or $plan.ExcelGroups.Count -or $plan.FastProjects.Count -or $plan.ProcessProjects.Count) {
                throw 'Setup-only npm selection must not require Excel or runtime validation.'
            }
            if (($plan.CiTestGroups -join ',') -ne 'Tooling') { throw 'Wrong setup validation group.' }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task CliAcceptanceChanges_RequireRuntimeAndAcceptanceValidation()
    {
        var result = await RunAsync("""
            $plan = Get-ValidationPlan -Paths 'tests\ExcelMcp.CLI.Tests\Integration\CliWorkflowAcceptanceTests.cs'
            if (-not $plan.Excel -or -not $plan.Build -or -not $plan.SourceChecks) { throw 'Runtime/E2E validation missing.' }
            if ('Acceptance' -notin $plan.ExcelGroups) { throw 'Acceptance group missing.' }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task MixedPaths_UnionGroupsAndNormalizeWindowsPaths()
    {
        var result = await RunAsync("""
            $plan = Get-ValidationPlan -Paths @('src\ExcelMcp.Core\Commands\Range\Deleted.cs', 'src/ExcelMcp.Core/Commands/PowerQuery/Renamed.cs')
            if (($plan.ExcelGroups -join ',') -ne 'Acceptance,Data,Editing') { throw "Wrong union: $($plan.ExcelGroups)" }
            if (($plan.FastProjects -join ',') -ne 'CLI,ComInterop,Core,McpServer,Service') { throw 'Shared runtime coverage missing.' }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task FullSelection_ContainsEveryHostedGroupAndPackage()
    {
        var result = await RunAsync("""
            $plan = Get-ValidationPlan -Full
            if (($plan.CiTestGroups -join ',') -ne 'Fast,Process,Tooling') { throw 'Full groups missing.' }
            foreach ($project in $plan.ToolingProjects) {
                if ($plan.ToolingFilters.$project -ne 'RequiresExcel=false') { throw 'Full tooling selection narrowed.' }
            }
            foreach ($component in @('Cli','Mcp','Extension','Mcpb','Skills','Plugins')) {
                if (-not $plan.$component) { throw "Missing $component package." }
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
        var info = new ProcessStartInfo("pwsh")
        {
            WorkingDirectory = root.FullName,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        if (environmentVariables is not null)
        {
            foreach (var (name, value) in environmentVariables) { info.Environment[name] = value; }
        }
        foreach (var name in info.Environment.Keys
            .Where(name => name.StartsWith("GIT_", StringComparison.OrdinalIgnoreCase)).ToArray())
        {
            info.Environment.Remove(name);
        }
        foreach (var argument in new[] { "-NoProfile", "-Command", $"$ErrorActionPreference='Stop'; . '{script}'; {command}" })
        {
            info.ArgumentList.Add(argument);
        }
        using var process = Process.Start(info);
        Assert.NotNull(process);
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try
        {
            await process.WaitForExitAsync(deadline.Token);
        }
        catch (OperationCanceledException) when (deadline.IsCancellationRequested)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw new TimeoutException($"Selection regression exceeded its deadline.\n{await stdout}\n{await stderr}");
        }
        return (process.ExitCode, await stdout + await stderr);
    }
}
