using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ValidationOrchestrationTests
{
    [Fact]
    public async Task ChildCommands_IgnoreAmbientGitRepositoryAndIndex()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $ambient = @(Get-ChildItem Env: | Where-Object Name -like 'GIT_*')
            if ($ambient.Count) { throw "Inherited Git context: $($ambient.Name -join ', ')" }
            if ($env:EXCELMCP_FIXTURE_MARKER -ne 'retained') { throw 'Non-Git environment was removed.' }
            """, new Dictionary<string, string>
        {
            ["GIT_DIR"] = "fixture-parent-git",
            ["GIT_COMMON_DIR"] = "fixture-parent-common",
            ["GIT_WORK_TREE"] = "fixture-parent-worktree",
            ["GIT_INDEX_FILE"] = "fixture-parent-index",
            ["GIT_PREFIX"] = "fixture-prefix",
            ["GIT_CONFIG_COUNT"] = "1",
            ["GIT_CONFIG_KEY_0"] = "core.bare",
            ["GIT_CONFIG_VALUE_0"] = "true",
            ["GIT_CONFIG_PARAMETERS"] = "'core.bare=true'",
            ["EXCELMCP_FIXTURE_MARKER"] = "retained"
        });
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData("success", "false", "skipped", true)]
    [InlineData("success", "true", "success", true)]
    [InlineData("failure", "false", "skipped", false)]
    [InlineData("cancelled", "false", "skipped", false)]
    [InlineData("success", "true", "failure", false)]
    [InlineData("success", "true", "cancelled", false)]
    [InlineData("success", "true", "skipped", false)]
    [InlineData("success", "false", "success", false)]
    [InlineData("success", "", "skipped", false)]
    public async Task Completion_RequiresDetectionAndExactSelectedResult(
        string detection, string selected, string result, bool succeeds)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            & .\scripts\Test-CiCompletion.ps1 -Detection '{{detection}}' -Tests '{{result}}' `
                -Packages skipped -Npm skipped -Lockfiles skipped -SelectedTests '{{selected}}' `
                -SelectedPackages false -SelectedNpm false -SelectedLockfiles false
            """);
        Assert.Equal(succeeds, run.ExitCode == 0);
        if (succeeds) { Assert.Contains("All selected CI checks succeeded.", run.Output, StringComparison.Ordinal); }
    }

    [Theory]
    [InlineData("Completed", 3, 3, true)]
    [InlineData("Completed", 0, 0, false)]
    [InlineData("Completed", 3, 2, false)]
    [InlineData("Failed", 3, 3, false)]
    public async Task Reports_RejectEmptySkippedAndAssemblyCleanupFailures(
        string outcome, int total, int passed, bool succeeds)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            . .\scripts\Invoke-TestStage.ps1
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Report.$([Guid]::NewGuid().ToString('N')).trx"
            try {
                '<TestRun xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010"><ResultSummary outcome="{{outcome}}"><Counters total="{{total}}" passed="{{passed}}"/></ResultSummary></TestRun>' |
                    Set-Content -LiteralPath $file
                Assert-TestReport -Path $file
            } finally { Remove-Item -LiteralPath $file }
            """);
        Assert.Equal(succeeds, run.ExitCode == 0);
    }

    [Fact]
    public async Task MissingReport_IsNeverSuccess()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\Invoke-TestStage.ps1
            Assert-TestReport -Path (Join-Path ([IO.Path]::GetTempPath()) "Missing-$([Guid]::NewGuid().ToString('N')).trx")
            """);
        Assert.NotEqual(0, run.ExitCode);
        Assert.Contains("Missing test report", run.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UnselectedGroup_IsRejectedBeforeTestExecution()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Plan.$([Guid]::NewGuid().ToString('N')).json"
            try {
                Get-ValidationPlan -Paths README.md | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $file
                & .\scripts\Invoke-ExcelFreeTests.ps1 -PlanFile $file -Group Process
            } finally { Remove-Item -LiteralPath $file }
            """);
        Assert.NotEqual(0, run.ExitCode);
        Assert.Contains("was not selected", run.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task HardDeadline_TerminatesTheStartedChildAndCannotPass()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\Invoke-TestStage.ps1
            $sandbox = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Deadline.$([Guid]::NewGuid().ToString('N'))"
            New-Item -ItemType Directory -Path $sandbox | Out-Null
            $project = Join-Path $sandbox fixture.proj
            $child = Join-Path $sandbox child.ps1
            $identity = Join-Path $sandbox child.json
            try {
                '$ErrorActionPreference="Stop"; $process = Get-Process -Id $PID; @{ pid=$PID; started=$process.StartTime.ToUniversalTime().Ticks } | ConvertTo-Json | Set-Content -LiteralPath "' + $identity + '"; Start-Sleep -Seconds 60' |
                    Set-Content -LiteralPath $child
                $escaped = [Security.SecurityElement]::Escape($child)
                '<Project><Target Name="VSTest"><Exec Command="pwsh -NoProfile -File &quot;' + $escaped + '&quot;" /></Target></Project>' |
                    Set-Content -LiteralPath $project
                try {
                    Invoke-TestStage -Project $project -Filter fixture -ResultsDirectory $sandbox -Name deadline -DeadlineSeconds 10
                    throw 'Deadline returned success.'
                } catch {
                    if ($_.Exception.Message -notmatch 'exceeded the 10-second hard deadline') { throw }
                }
                if (-not (Test-Path -LiteralPath $identity)) { throw 'The child never started.' }
                $owned = Get-Content -LiteralPath $identity -Raw | ConvertFrom-Json
                $survivor = Get-Process -Id $owned.pid -ErrorAction SilentlyContinue
                if ($survivor -and $survivor.StartTime.ToUniversalTime().Ticks -eq $owned.started) {
                    $survivor.Kill($true)
                    $survivor.WaitForExit()
                    throw 'The exact started child survived its hard deadline.'
                }
            } finally {
                Remove-Item -LiteralPath $sandbox -Recurse -Force
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task GitComparison_UsesMergeBaseAndFailsOnInvalidRevisions()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $script = Join-Path (Get-Location) 'scripts\Get-CiValidationPlan.ps1'
            $sandbox = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.GitSelection.$([Guid]::NewGuid().ToString('N'))"
            New-Item -ItemType Directory -Path $sandbox | Out-Null
            Push-Location $sandbox
            function Commit-Fixture {
                git add --all
                if ($LASTEXITCODE -ne 0) { throw 'Fixture stage failed.' }
                git -c user.name=Fixture -c user.email=fixture@example.invalid commit --quiet -m fixture
                if ($LASTEXITCODE -ne 0) { throw 'Fixture commit failed.' }
            }
            try {
                git init --quiet -b baseline
                if ($LASTEXITCODE -ne 0) { throw 'Fixture initialization failed.' }
                Set-Content README.md initial
                Commit-Fixture
                git switch --quiet -c feature
                if ($LASTEXITCODE -ne 0) { throw 'Fixture branch failed.' }
                Set-Content README.md changed
                Commit-Fixture
                $head = git rev-parse HEAD
                git switch --quiet baseline
                if ($LASTEXITCODE -ne 0) { throw 'Fixture branch failed.' }
                New-Item -ItemType Directory src\ExcelMcp.Service | Out-Null
                Set-Content src\ExcelMcp.Service\Service.cs changed
                Commit-Fixture
                $base = git rev-parse HEAD
                & $script -BaseRef $base -HeadRef $head -OutputPath merge.json
                $merge = Get-Content merge.json -Raw | ConvertFrom-Json
                if ($merge.CiTestGroups.Count) { throw 'PR selection included unrelated base-branch work.' }
                & $script -BaseRef $base -HeadRef $head -Comparison Direct -OutputPath direct.json
                $direct = Get-Content direct.json -Raw | ConvertFrom-Json
                if (-not $direct.Packages) { throw 'Direct range missed deleted runtime input.' }
                try {
                    & $script -BaseRef invalid -HeadRef $head -OutputPath invalid.json
                    throw 'Invalid revisions were accepted.'
                } catch {
                    if ($_.Exception.Message -notmatch 'Cannot determine changed validation inputs') { throw }
                }
            } finally {
                Pop-Location
                Remove-Item -LiteralPath $sandbox -Recurse -Force
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }
}
