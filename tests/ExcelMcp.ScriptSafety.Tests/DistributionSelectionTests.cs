using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class DistributionSelectionTests
{
    [Fact]
    public async Task ChangedInputs_SelectOnlyDependentDistributions()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $cases = @(
                @{ Path='gh-pages/sitegen/analytics.py'; Components=''; Npm=$false },
                @{ Path='docs/ARCHITECTURE.md'; Components=''; Npm=$false },
                @{ Path='src/ExcelMcp.Core/Commands/Range/RangeCommands.Values.cs'; Components=''; Npm=$false },
                @{ Path='tests/ExcelMcp.Packaging.Tests/PluginBuildTests.cs'; Components=''; Npm=$false },
                @{ Path='vscode-extension/package-lock.json'; Components='Extension'; Npm=$false },
                @{ Path='skills/excel-cli/SKILL.md'; Components='Plugins'; Npm=$false },
                @{ Path='skills/excel-cli-report-formatting/SKILL.md'; Components='Plugins'; Npm=$false },
                @{ Path='skills/excel-mcp-report-formatting/SKILL.md'; Components='Extension,Plugins'; Npm=$false },
                @{ Path='docs/reference/report-formatting.md'; Components='Extension,Plugins'; Npm=$false },
                @{ Path='.github/plugins/excel-cli/plugin.json'; Components='Plugins'; Npm=$false },
                @{ Path='.github/plugins/marketplace-repo/README.md'; Components='Plugins'; Npm=$false },
                @{ Path='scripts/Build-Plugins.ps1'; Components='Plugins'; Npm=$false },
                @{ Path='scripts/Build-AgentSkills.ps1'; Components='Extension,Plugins'; Npm=$false },
                @{ Path='scripts/Publish-PreparedPlugins.ps1'; Components=''; Npm=$false },
                @{ Path='mcpb/manifest.json'; Components='Mcpb'; Npm=$false },
                @{ Path='npm-packages/excelcli/package.json'; Components='Cli'; Npm=$true },
                @{ Path='npm-packages/mcp-server-excel/package.json'; Components='Mcp'; Npm=$true },
                @{ Path='npm-packages/shared/launcher.js'; Components='Cli,Mcp'; Npm=$true },
                @{ Path='npm-packages/shared/test/launcher.test.js'; Components=''; Npm=$true },
                @{ Path='scripts/Build-NpmPackages.ps1'; Components='Cli,Mcp'; Npm=$true },
                @{ Path='scripts/ExtensionPackaging.ps1'; Components='Extension'; Npm=$false },
                @{ Path='scripts/PackageHelpers.ps1'; Components='Cli,Extension,Mcp,Mcpb,Plugins'; Npm=$true },
                @{ Path='src/ExcelMcp.CLI/ExcelMcp.CLI.csproj'; Components='Cli'; Npm=$false },
                @{ Path='src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj'; Components='Extension,Mcp'; Npm=$false },
                @{ Path='src/ExcelMcp.Core/ExcelMcp.Core.csproj'; Components='Cli,Extension,Mcp'; Npm=$false },
                @{ Path='global.json'; Components='Cli,Extension,Mcp,Mcpb,Plugins'; Npm=$true },
                @{ Path='Directory.Build.props'; Components='Cli,Extension,Mcp,Mcpb,Plugins'; Npm=$true },
                @{ Path='Directory.Build.targets'; Components='Cli,Extension,Mcp,Mcpb,Plugins'; Npm=$true },
                @{ Path='Directory.Packages.props'; Components='Cli,Extension,Mcp,Mcpb,Plugins'; Npm=$true },
                @{ Path='NuGet.Config'; Components='Cli,Extension,Mcp,Mcpb,Plugins'; Npm=$true },
                @{ Path='package.json'; Components='Cli,Extension,Mcp,Mcpb,Plugins'; Npm=$true },
                @{ Path='README.md'; Components='Cli,Extension,Mcp,Mcpb,Plugins'; Npm=$true },
                @{ Path='LICENSE'; Components='Cli,Extension,Mcp,Mcpb,Plugins'; Npm=$true },
                @{ Path='.github/workflows/release.yml'; Components='Cli,Extension,Mcp,Mcpb,Plugins'; Npm=$true },
                @{ Path='.github/workflows/ci.yml'; Components='Cli,Extension,Mcp,Mcpb,Plugins'; Npm=$true }
            )
            foreach ($case in $cases) {
                $plan = Get-ValidationPlan -Paths $case.Path
                if (($plan.PackageComponents -join ',') -cne $case.Components -or $plan.NpmTests -ne $case.Npm) {
                    throw "Wrong distributions for $($case.Path): $($plan.PackageComponents), npm=$($plan.NpmTests)"
                }
            }
            $mixed = Get-ValidationPlan -Paths @('vscode-extension\src\extension.ts','npm-packages/excelcli/bin/excelcli.js')
            if (($mixed.PackageComponents -join ',') -cne 'Cli,Extension' -or -not $mixed.NpmTests) { throw 'Mixed dependencies lost.' }
            $full = Get-ValidationPlan -Full
            if (($full.PackageComponents -join ',') -cne 'Cli,Extension,Mcp,Mcpb,Plugins' -or -not $full.NpmTests) { throw 'Full distributions missing.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task CiOutputs_PreserveJobAndSetupSelections()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $directory = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.DistributionOutputs.$([Guid]::NewGuid().ToString('N'))"
            New-Item -ItemType Directory -Path $directory | Out-Null
            $previousOutput = $env:GITHUB_OUTPUT
            try {
                foreach ($case in @(
                    @{ Path='docs/guide.md'; Packages='false'; Npm='false'; Skills='false'; Dotnet='false' },
                    @{ Path='mcpb/manifest.json'; Packages='true'; Npm='false'; Skills='false'; Dotnet='false' },
                    @{ Path='.github/plugins/excel-cli/plugin.json'; Packages='true'; Npm='false'; Skills='true'; Dotnet='false' },
                    @{ Path='vscode-extension/package.json'; Packages='true'; Npm='false'; Skills='true'; Dotnet='true' },
                    @{ Path='npm-packages/excelcli/package.json'; Packages='true'; Npm='true'; Skills='false'; Dotnet='true' },
                    @{ Path='npm-packages/shared/test/launcher.test.js'; Packages='false'; Npm='true'; Skills='false'; Dotnet='false' },
                    @{ Path='all'; Packages='true'; Npm='true'; Skills='true'; Dotnet='true' }
                )) {
                    function git { $case.Path; $global:LASTEXITCODE = 0 }
                    $env:GITHUB_OUTPUT = Join-Path $directory outputs.txt
                    Set-Content $env:GITHUB_OUTPUT ''
                    & .\scripts\Get-CiValidationPlan.ps1 -BaseRef baseline -Full:($case.Path -eq 'all') -OutputPath (Join-Path $directory plan.json)
                    $outputs = Get-Content $env:GITHUB_OUTPUT
                    foreach ($expected in @(
                        "packages=$($case.Packages)", "npm=$($case.Npm)",
                        "package_skills=$($case.Skills)", "package_dotnet=$($case.Dotnet)"
                    )) {
                        if ($expected -cnotin $outputs) { throw "Missing $expected for $($case.Path)." }
                    }
                }
                $workflow = Get-Content .\.github\workflows\ci.yml -Raw
                foreach ($name in @('packages','npm','package_skills','package_dotnet')) {
                    if (-not $workflow.Contains("$name`: " + '${{ steps.select.outputs.' + $name + ' }}')) {
                        throw "Workflow does not export $name."
                    }
                }
                foreach ($name in @('Packages','Npm')) {
                    if (-not $workflow.Contains("SELECTED_$($name.ToUpperInvariant())`: " + '${{ needs.changes.outputs.' + $name.ToLowerInvariant() + ' }}') -or
                        -not $workflow.Contains("-Selected$name " + '$env:SELECTED_' + $name.ToUpperInvariant())) {
                        throw "Completion does not consume selected $name."
                    }
                }
            } finally {
                $env:GITHUB_OUTPUT = $previousOutput
                Remove-Item -LiteralPath $directory -Recurse -Force
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task SavedPlans_RejectMissingOrInvalidDistributionSelections()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\ValidationExecution.ps1
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Distributions.$([Guid]::NewGuid().ToString('N')).json"
            try {
                $plan = Get-ValidationPlan -Paths 'vscode-extension/src/extension.ts'
                $plan | ConvertTo-Json -Depth 10 | Set-Content $file
                $saved = Read-ValidationPlan $file
                if (($saved.PackageComponents -join ',') -cne 'Extension' -or $saved.NpmTests) { throw 'Saved distributions changed.' }
                foreach ($kind in @('missing-components','bad-component','duplicate-component','missing-npm','string-npm','old-schema')) {
                    $bad = Get-ValidationPlan -Paths 'vscode-extension/src/extension.ts'
                    switch ($kind) {
                        missing-components { $bad.Remove('PackageComponents') }
                        bad-component { $bad.PackageComponents = @('Unknown') }
                        duplicate-component { $bad.PackageComponents = @('Extension','Extension') }
                        missing-npm { $bad.Remove('NpmTests') }
                        string-npm { $bad.NpmTests = 'false' }
                        old-schema { $bad.SchemaVersion = 1 }
                    }
                    $bad | ConvertTo-Json -Depth 10 | Set-Content $file
                    $rejected = $false
                    try { $null = Read-ValidationPlan $file } catch { $rejected = $true }
                    if (-not $rejected) { throw "Invalid distribution plan accepted: $kind" }
                }
            } finally { Remove-Item -LiteralPath $file }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task Completion_RequiresExactPackageAndNpmOutcomes()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            foreach ($job in @('Packages','Npm')) {
                foreach ($selected in @('true','false','','invalid')) {
                    foreach ($result in @('success','skipped','failure','cancelled')) {
                        $arguments = @{
                            Detection='success'; Tests='skipped'; SelectedTests='false'
                            BuildOnly='skipped'; SelectedBuildOnly='false'
                            Packages='skipped'; SelectedPackages='false'; Npm='skipped'; SelectedNpm='false'
                        }
                        $arguments[$job] = $result
                        $arguments["Selected$job"] = $selected
                        $expected = ($selected -ceq 'true' -and $result -ceq 'success') -or
                            ($selected -ceq 'false' -and $result -ceq 'skipped')
                        $succeeded = $false
                        try { & .\scripts\Test-CiCompletion.ps1 @arguments; $succeeded = $true } catch { }
                        if ($succeeded -ne $expected) { throw "Wrong completion for $job selected='$selected' result='$result'." }
                    }
                }
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }
}
