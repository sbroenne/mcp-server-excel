using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ValidationOrchestrationTests
{
    [Fact]
    public async Task CliAcceptance_WatchdogAllowsNativeWorkflowDeadline()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $tokens = $null
            $errors = $null
            $ast = [Management.Automation.Language.Parser]::ParseFile(
                (Resolve-Path .\scripts\Test-E2E.ps1).Path, [ref]$tokens, [ref]$errors)
            if ($errors.Count) { throw 'E2E script does not parse.' }
            $selection = $ast.Find({
                param($node)
                $node -is [Management.Automation.Language.SwitchStatementAst] -and
                $node.Condition.Extent.Text -eq '$stage'
            }, $true)
            if (-not $selection) { throw 'Acceptance stage selection missing.' }
            $stage = 'Cli'
            $parameters = @{}
            & ([scriptblock]::Create($selection.Extent.Text))
            if ($parameters.DeadlineSeconds -ne 600) { throw 'CLI hard deadline changed.' }
            $watchdogSeconds = [int]::Parse($parameters.HangTimeout.TrimEnd('m')) * 60
            if ($watchdogSeconds -le 480 -or $watchdogSeconds -gt $parameters.DeadlineSeconds) {
                throw 'CLI watchdog must allow the eight-minute native workflow without extending the hard deadline.'
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task CiPreparatoryBuilds_DisableBuildServers()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $workflow = Get-Content .\.github\workflows\ci.yml -Raw
            if (-not $workflow.Contains('./scripts/Build-CiInputs.ps1')) { throw 'Scoped preparation was not wired.' }
            $builds = @(Get-Content .\scripts\Build-CiInputs.ps1 | Where-Object { $_ -match '& dotnet build' })
            if ($builds.Count -ne 1) { throw "Unexpected preparatory build count: $($builds.Count)" }
            foreach ($build in $builds) {
                if ($build -notmatch '--disable-build-servers') { throw 'Preparatory build can retain a locking build server.' }
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task CiSelection_ExportsAndUsesTheSourceCheckGroup()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $sandbox = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.CiOutputs.$([Guid]::NewGuid().ToString('N'))"
            New-Item -ItemType Directory -Path $sandbox | Out-Null
            $previousOutput = $env:GITHUB_OUTPUT
            try {
                $env:GITHUB_OUTPUT = Join-Path $sandbox outputs.txt
                & .\scripts\Get-CiValidationPlan.ps1 -Full -OutputPath (Join-Path $sandbox plan.json)
                if ('source_checks_group=Fast' -cnotin (Get-Content -LiteralPath $env:GITHUB_OUTPUT)) {
                    throw 'Source-check selection was not exported.'
                }
                $workflow = Get-Content .\.github\workflows\ci.yml -Raw
                if (-not $workflow.Contains('source_checks_group: ${{ steps.select.outputs.source_checks_group }}') -or
                    -not $workflow.Contains('if: matrix.group == needs.changes.outputs.source_checks_group')) {
                    throw 'Workflow does not route the selected source checks.'
                }
            } finally {
                $env:GITHUB_OUTPUT = $previousOutput
                Remove-Item -LiteralPath $sandbox -Recurse -Force
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData("tests/AGENTS.md", "", false)]
    [InlineData("scripts/PluginContent.mjs", "javascript-typescript", false)]
    [InlineData("src/ExcelMcp.McpServer/Program.cs", "csharp", true)]
    [InlineData("all", "actions,csharp,javascript-typescript,python", true)]
    public async Task CiSelection_ExportsLanguageMatrixAndPackageBuild(string path, string languages, bool packageBuild)
    {
        var run = await ValidationSelectionTests.RunAsync($$$"""
            $sandbox = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.LanguageOutputs.$([Guid]::NewGuid().ToString('N'))"
            New-Item -ItemType Directory -Path $sandbox | Out-Null
            $previousOutput = $env:GITHUB_OUTPUT
            try {
                $env:GITHUB_OUTPUT = Join-Path $sandbox outputs.txt
                if ('{{{path}}}' -ne 'all') {
                    function git { '{{{path}}}'; $global:LASTEXITCODE = 0 }
                }
                & .\scripts\Get-CiValidationPlan.ps1 -BaseRef baseline -Full:('{{{path}}}' -eq 'all') `
                    -OutputPath (Join-Path $sandbox plan.json)
                $outputs = @{}
                foreach ($line in Get-Content -LiteralPath $env:GITHUB_OUTPUT) {
                    $key, $value = $line -split '=', 2
                    $outputs[$key] = $value
                }
                $matrix = $outputs.codeql_matrix | ConvertFrom-Json
                if (($matrix.include.language -join ',') -cne '{{{languages}}}') { throw 'Wrong exported language matrix.' }
                $expectedCodeQl = if ('{{{languages}}}') { 'true' } else { 'false' }
                if ($outputs.codeql -cne $expectedCodeQl -or
                    $outputs.package_build -cne '{{{packageBuild.ToString().ToLowerInvariant()}}}') {
                    throw 'Wrong selection output flags.'
                }
                foreach ($entry in $matrix.include) {
                    $os = if ($entry.language -eq 'csharp') { 'windows-latest' } else { 'ubuntu-26.04' }
                    $mode = if ($entry.language -eq 'csharp') { 'manual' } else { 'none' }
                    if ($entry.os -cne $os -or $entry.'build-mode' -cne $mode) { throw 'Language extraction changed.' }
                }
                $workflow = Get-Content .\.github\workflows\codeql.yml -Raw
                foreach ($required in @(
                    './scripts/Get-CiValidationPlan.ps1',
                    'matrix: ${{ steps.select.outputs.codeql_matrix }}',
                    'has-any: ${{ steps.select.outputs.codeql }}',
                    'fetch-depth: 0',
                    '-Full:($env:EVENT -ne ''pull_request'')',
                    'dotnet build Sbroenne.ExcelMcp.sln --no-restore --configuration Release --no-incremental',
                    'name: CodeQL Completion',
                    '[ "$DETECTION" = success ]',
                    'false:skipped)')) {
                    if (-not $workflow.Contains($required)) { throw "Missing CodeQL guarantee: $required" }
                }
            } finally {
                $env:GITHUB_OUTPUT = $previousOutput
                Remove-Item -LiteralPath $sandbox -Recurse -Force
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

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
                -Packages skipped -Npm skipped -SelectedTests '{{selected}}' `
                -SelectedPackages false -SelectedNpm false
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
                $rows = (1..{{total}} | ForEach-Object {
                    if ({{total}} -gt 0) { '<UnitTestResult testName="case" outcome="Passed"/>' }
                }) -join ''
                '<TestRun xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010"><Results>' + $rows +
                    '</Results><ResultSummary outcome="{{outcome}}"><Counters total="{{total}}" executed="{{total}}" passed="{{passed}}" failed="0" notExecuted="0"/></ResultSummary></TestRun>' |
                    Set-Content -LiteralPath $file
                Assert-TestReport -Path $file
            } finally { Remove-Item -LiteralPath $file }
            """);
        Assert.Equal(succeeds, run.ExitCode == 0);
    }

    [Theory]
    [InlineData("row")]
    [InlineData("missing-row")]
    [InlineData("failed")]
    [InlineData("notExecuted")]
    [InlineData("executed")]
    [InlineData("timeout")]
    public async Task Reports_RejectContradictoryResults(string change)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            . .\scripts\Invoke-TestStage.ps1
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Report.$([Guid]::NewGuid().ToString('N')).trx"
            try {
                [xml]$report = '<TestRun><Results><UnitTestResult testName="case" outcome="Passed"/></Results><ResultSummary outcome="Completed"><Counters total="1" executed="1" passed="1" failed="0" notExecuted="0"/></ResultSummary></TestRun>'
                switch ('{{change}}') {
                    row { $report.TestRun.Results.UnitTestResult.outcome = 'Failed' }
                    missing-row { [void]$report.TestRun.Results.RemoveAll() }
                    executed { $report.TestRun.ResultSummary.Counters.executed = '0' }
                    default { $report.TestRun.ResultSummary.Counters.SetAttribute('{{change}}', '1') }
                }
                $report.Save($file)
                Assert-TestReport -Path $file
            } finally { Remove-Item -LiteralPath $file }
            """);
        Assert.NotEqual(0, run.ExitCode);
    }

    [Theory]
    [InlineData("@('case(row: 1)', 'case(row: 1)')", "@('case(row: 1)', 'case(row: 1)')", true)]
    [InlineData("@('case(row: 1)', 'other')", "@('other', 'case(row: 1)')", true)]
    [InlineData("@('case(row: 1)', 'case(row: 1)')", "@('case(row: 1)')", false)]
    [InlineData("@('case(row: 1)', 'other')", "@('case(row: 1)', 'case(row: 1)')", false)]
    [InlineData("@('case')", "@('case', 'extra')", false)]
    [InlineData("@('Case')", "@('case')", false)]
    [InlineData("@()", "@('case')", false)]
    public async Task Reports_ReconcileDiscoveredCasesIncludingTheoryMultiplicity(
        string expected, string actual, bool succeeds)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            . .\scripts\Invoke-TestStage.ps1
            $file = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Reconciliation.$([Guid]::NewGuid().ToString('N')).trx"
            try {
                $cases = {{actual}}
                $rows = ($cases | ForEach-Object { '<UnitTestResult testName="' + $_ + '" outcome="Passed"/>' }) -join ''
                '<TestRun><Results>' + $rows + '</Results><ResultSummary outcome="Completed"><Counters total="' +
                    $cases.Count + '" executed="' + $cases.Count + '" passed="' + $cases.Count +
                    '" failed="0" notExecuted="0"/></ResultSummary></TestRun>' | Set-Content -LiteralPath $file
                Assert-TestReport -Path $file -ExpectedCases {{expected}}
            } finally { Remove-Item -LiteralPath $file }
            """);
        Assert.Equal(succeeds, run.ExitCode == 0);
        if (!succeeds)
        {
            Assert.Contains("Test case reconciliation failed", run.Output, StringComparison.Ordinal);
        }
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
                if ($merge.FastProjects.Count -or $merge.ProcessProjects.Count -or
                    ($merge.CiTestGroups -join ',') -cne 'Tooling' -or $merge.Excel) {
                    throw 'PR selection included unrelated base-branch work.'
                }
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
