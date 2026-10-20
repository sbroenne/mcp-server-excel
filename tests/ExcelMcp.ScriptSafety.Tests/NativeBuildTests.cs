using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class NativeBuildTests
{
    [Theory]
    [InlineData("empty")]
    [InlineData("owner")]
    [InlineData("filter")]
    public async Task SavedExcelSelections_CannotBecomeEmptyOrChangeOwnership(string invalid)
    {
        var result = await ValidationSelectionTests.RunAsync($$"""
            . .\scripts\ValidationExecution.ps1
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.InvalidPlan.$([Guid]::NewGuid().ToString('N')).json"
            try {
                $plan = Get-ValidationPlan -Paths 'src/ExcelMcp.Core/Commands/Range/Commands.cs'
                switch ('{{invalid}}') {
                    empty { $plan.ExcelSelections = @() }
                    owner { $plan.ExcelSelections[0].Project = 'Unknown' }
                    filter { $plan.ExcelSelections[0].Filter = '' }
                }
                $plan | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $file
                try { Read-ValidationPlan $file; throw 'Invalid saved plan was accepted.' }
                catch { if ($_ -notmatch 'no test filters|Invalid saved Excel selection') { throw } }
            } finally { Remove-Item -LiteralPath $file -Force }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task TestCommand_BuildsByDefaultAndRejectsZeroDeadlines()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            . .\scripts\ValidationExecution.ps1
            $script = Get-Content .\build.ps1 -Raw
            $script = $script.Replace(". (Join-Path `$PSScriptRoot 'scripts\ValidationExecution.ps1')", '')
            $command = [scriptblock]::Create($script)
            function Invoke-SelectedBuild { $global:buildCalled = $true }
            function Invoke-SelectedFreeTests {}
            foreach ($noBuild in @($false, $true)) {
                $global:buildCalled = $false
                $arguments = @('test','--paths-json','["README.md"]','--results-directory','TestResults\command-contract')
                if ($noBuild) { $arguments += '--no-build' }
                & $command -Arguments $arguments
                if ($global:buildCalled -eq $noBuild) { throw 'Test command build default changed.' }
            }
            try { & $command -Arguments @('test','--paths-json','["README.md"]','--deadline-seconds','0'); throw 'Zero deadline accepted.' }
            catch { if ($_ -notmatch '1-28800 seconds') { throw } }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task Deadline_StopsOwnedChildrenEvenAfterTheirParentExits()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            . .\scripts\Invoke-NativeCommand.ps1
            $directory = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Orphan.$([Guid]::NewGuid().ToString('N'))"
            New-Item -ItemType Directory $directory | Out-Null
            $identity = Join-Path $directory child.json
            try {
                $child = Join-Path $directory child.ps1
                '$p=Get-Process -Id $PID; @{Id=$PID; Start=$p.StartTime.ToUniversalTime().Ticks} | ConvertTo-Json | Set-Content -LiteralPath "' + $identity + '"; Start-Sleep 60' |
                    Set-Content $child
                $parent = Join-Path $directory parent.ps1
                '[Diagnostics.Process]::Start("pwsh", ''-NoProfile -File "' + $child + '"'') | Out-Null' |
                    Set-Content $parent
                try {
                    Invoke-NativeCommand -Executable pwsh -Arguments @('-NoProfile','-File',$parent) -DeadlineSeconds 10
                    throw 'An unfinished process tree returned success.'
                } catch {
                    if ($_ -notmatch 'hard deadline') { throw }
                }
                if (-not (Test-Path $identity)) { throw 'The child never started.' }
                $owned = Get-Content $identity -Raw | ConvertFrom-Json
                $process = Get-Process -Id $owned.Id -ErrorAction SilentlyContinue
                if ($process -and $process.StartTime.ToUniversalTime().Ticks -eq $owned.Start) {
                    $process.Kill($true)
                    $process.WaitForExit()
                    throw 'The owned child survived.'
                }
            } finally { Remove-Item -LiteralPath $directory -Recurse -Force }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task HelpAndSelection_DoNotBuildOrLoadACustomTool()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            function Get-Command { throw 'A custom SDK bootstrap was attempted.' }
            $help = & .\build.ps1 --help
            if (($help -join "`n") -notmatch 'ExcelMcp') { throw 'Missing help.' }
            $plan = Get-ValidationPlan -Paths 'src/ExcelMcp.Core/Commands/Range/RangeCommands.Values.cs'
            if ($plan.FastFilters.Core -notmatch 'Feature=Range') { throw 'Native test labels were not selected.' }
            if (-not $plan.ExcelSelections.Count) { throw 'Workbook coverage was lost.' }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task StrictSettings_EnableEmptySkippedAndWarningFailures()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            $settings = [xml](Get-Content .\tests\strict.runsettings -Raw)
            if ($settings.RunSettings.RunConfiguration.TreatNoTestsAsError -cne 'true' -or
                $settings.RunSettings.xUnit.FailSkips -cne 'true' -or
                $settings.RunSettings.xUnit.FailWarns -cne 'true') {
                throw 'Native runner failure settings are missing.'
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task SharedSourceSelection_IsExplicitAndRejectsUnmappedCommandAreas()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            $shared = Get-ValidationPlan -Paths 'src/ExcelMcp.Core/Helpers/ComInteropHelpers.cs'
            if (-not $shared.Excel -or -not $shared.Build) { throw 'Shared consumers missing.' }
            try {
                Get-ValidationPlan -Paths 'src/ExcelMcp.Core/Commands/Unmapped/Removed.cs'
                throw 'Unknown operational input was accepted.'
            } catch {
                if ($_ -notmatch 'No validation mapping') { throw }
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }
}
