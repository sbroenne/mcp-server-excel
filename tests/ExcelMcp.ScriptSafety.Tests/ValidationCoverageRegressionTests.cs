using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ValidationCoverageRegressionTests
{
    [Theory]
    [InlineData("infrastructure/azure/configure-analytics-oidc.ps1")]
    [InlineData("infrastructure/azure/deploy-appinsights.ps1")]
    public async Task ExplicitOnDemandCases_SurviveSavedPlansAndBothExecutionPaths(string path)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            . .\scripts\ValidationExecution.ps1
            $plan = Get-ValidationPlan -Paths '{{path}}'
            $saved = $plan | ConvertTo-Json -Depth 10 | ConvertFrom-Json -AsHashtable
            foreach ($parameters in @(@{ Local=$true }, @{ Group='Tooling' })) {
                $selected = @(Get-FreeTestSelection -Plan $saved @parameters)
                if ($selected.Count -ne 1 -or $selected[0].Owner -ne 'ScriptSafety' -or $saved.Excel) { throw 'Wrong OnDemand owner.' }
                foreach ($required in @('RequiresExcel=false','RunType!=OnDemand','RunType=OnDemand',
                    'AzurePreview_NeverMutatesAndRejectsFailedReads','AnalyticsOidc_SwitchesToSelectedTenantAndRestoresOriginalSubscription')) {
                    if (-not $selected[0].Filter.Contains($required)) { throw "Selected Azure case lost: $required" }
                }
            }
            $ordinary = Get-ValidationPlan -Paths 'videos/excel-mcp-intro/Capture-Evidence.ps1'
            $selected = @(Get-FreeTestSelection -Plan $ordinary -Group Tooling)
            if ($selected[0].Filter.Contains('RunType=OnDemand')) { throw 'Ordinary automation selected OnDemand cases.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData(false, "FullyQualifiedName=Selected.Case")]
    [InlineData(true, "")]
    public async Task SavedOnDemandFilters_RequireAnOwnerAndNonemptySelection(bool owner, string filter)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            . .\scripts\ValidationExecution.ps1
            $plan = Get-ValidationPlan -Paths '{{(owner ? "scripts/pre-commit.ps1" : "scripts/Build-Plugins.ps1")}}'
            $plan.ToolingOnDemandFilters.ScriptSafety = '{{filter}}'
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.InvalidPlan.$([Guid]::NewGuid().ToString('N')).json"
            try {
                $plan | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $file
                Read-ValidationPlan $file
            } finally { Remove-Item -LiteralPath $file }
            """);
        Assert.NotEqual(0, run.ExitCode);
        Assert.Contains("Selected OnDemand cases require", run.Output, StringComparison.Ordinal);
    }
}
