using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedValidationPolicyTests
{
    internal static readonly string Root = FindRoot();

    [Theory]
    [InlineData("Range", "RangeCommands.Values.cs", "Core")]
    [InlineData("PowerQuery", "PowerQueryCommands.Read.cs", "Core,Service")]
    public async Task FeatureEdits_SelectBehaviorWithoutUnchangedAdapters(string area, string file, string owners)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths 'src/ExcelMcp.Core/Commands/{{area}}/{{file}}'
            if (($plan.FastProjects -join ',') -cne '{{owners}}' -or $plan.ProcessProjects.Count -or $plan.ToolingProjects.Count) {
                throw 'Unchanged adapters or tooling selected.'
            }
            if (-not @($plan.ExcelSelections | Where-Object { $_.Area -eq '{{area}}' }).Count -or $plan.FullE2E) {
                throw 'Owning workbook selection missing or acceptance broadened.'
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData("src/ExcelMcp.Analyzers/SafetyAnalyzer.cs")]
    [InlineData("BannedSymbols.txt")]
    [InlineData("BannedSymbols.Mcp.txt")]
    public async Task CompilerRuleChanges_BuildEveryConsumerWithoutWorkbookTests(string path)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if (-not $plan.FullSolutionBuild -or $plan.Excel -or $plan.DocumentationCounts -or
                $plan.ToolingFilters.ScriptSafety -notmatch 'AnalyzerTests') { throw 'Compiler-rule consumers missing.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData("src/Shared/TelemetryUserIdentity.cs")]
    [InlineData("src/Shared/UsageTelemetry.cs")]
    [InlineData("src/Shared/UsageRecord.cs")]
    public async Task SharedTelemetry_SelectsBothEntryPointsButNoWorkbookTests(string path)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if (($plan.FastProjects -join ',') -cne 'CLI,McpServer' -or $plan.Excel -or $plan.FullE2E) {
                throw 'Wrong shared telemetry consumers.'
            }
            if ($plan.FastFilters.CLI -notmatch 'EntryPointTelemetryTests' -or
                $plan.FastFilters.McpServer -notmatch 'TelemetryUserIdentityTests') { throw 'Identity consumers missing.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData("src/ExcelMcp.CLI/Infrastructure/CliServiceTray.cs", "Process", "CliDaemonLifecycleTests")]
    [InlineData("src/ExcelMcp.CLI/Commands/BatchCommand.cs", "Fast", "Feature=Batch")]
    public async Task SpecialAdapterBoundaries_KeepTheirOwningCases(string path, string kind, string filter)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if (($plan.CiTestGroups -join ',') -cne '{{kind}}' -or
                $plan['{{kind}}Filters'].CLI -notmatch '{{filter}}' -or $plan.Excel) { throw 'Adapter boundary lost.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task AllCommandAreas_HaveExplicitRulesAndSavedPlansKeepTheirSelection()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            foreach ($directory in Get-ChildItem .\src\ExcelMcp.Core\Commands -Directory) {
                $plan = Get-ValidationPlan -Paths "src/ExcelMcp.Core/Commands/$($directory.Name)/Commands.cs"
                if (-not $plan.Build -or -not $plan.Reasons.Count) { throw "Missing mapping: $directory" }
            }
            $paths = @('src/ExcelMcp.Core/Commands/Range/Deleted.cs', 'src/ExcelMcp.Core/Commands/PowerQuery/Renamed.cs')
            $first = Get-ValidationPlan -Paths $paths | ConvertTo-Json -Depth 10
            [array]::Reverse($paths)
            $second = Get-ValidationPlan -Paths $paths | ConvertTo-Json -Depth 10
            if ($first -cne $second) { throw 'Selection depends on input order.' }
            $directory = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.PlanRoundTrip.$([Guid]::NewGuid().ToString('N'))"
            New-Item -ItemType Directory $directory | Out-Null
            try {
                $file = Join-Path $directory plan.json
                $first | Set-Content -LiteralPath $file
                . .\scripts\ValidationExecution.ps1
                $saved = Read-ValidationPlan $file
                if (($saved | ConvertTo-Json -Depth 10) -cne $first) { throw 'Saved selection changed.' }
            } finally { Remove-Item -LiteralPath $directory -Recurse -Force }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData("unknown-build-input.config")]
    [InlineData("src/ExcelMcp.Core/Commands/Unknown/Commands.cs")]
    [InlineData("tests/ExcelMcp.Unknown.Tests/Cases.cs")]
    [InlineData("../Directory.Build.props")]
    [InlineData("src/../../Directory.Build.props")]
    [InlineData("C:\\outside.cs")]
    [InlineData("")]
    public async Task UnknownOrInvalidPaths_FailInsteadOfSelectingEverything(string path)
    {
        var run = await ValidationSelectionTests.RunAsync($"Get-ValidationPlan -Paths '{path}'");
        Assert.NotEqual(0, run.ExitCode);
        Assert.Matches("No validation mapping|repository-relative|cannot be empty", run.Output);
    }

    [Theory]
    [InlineData("scripts/Test-E2E.ps1")]
    [InlineData("scripts/Test-CliWorkflow.ps1")]
    [InlineData("scripts/Stop-ExcelCliService.ps1")]
    public async Task AcceptanceBoundaryChanges_KeepBothRequiredEntryPoints(string path)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if (-not $plan.FullE2E -or -not $plan.Excel -or -not $plan.Build -or $plan.FullSolutionBuild) {
                throw 'Required acceptance boundary changed.'
            }
            if (($plan.ExcelSelections.Project -join ',') -cne 'CLI,McpServer') { throw 'An entry point was lost.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
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
