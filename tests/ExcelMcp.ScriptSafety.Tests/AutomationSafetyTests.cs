using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "AutomationSafety")]
public sealed partial class AutomationSafetyTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RunnerGuestCommands_RequirePowerShell7AndPreserveNativeFailure(bool fail)
    {
        var root = NewSandbox();
        try
        {
            var result = await RunAsync(root, $$"""
                . '{{Quote(Path.Combine(RepoRoot, "scripts", "AzureRunnerHost.ps1"))}}'
                New-Item -ItemType Directory -Path '{{Quote(Path.Combine(root, "PowerShell"))}}' | Out-Null
                New-Item -ItemType Junction -Path '{{Quote(Path.Combine(root, "PowerShell", "7"))}}' -Target $PSHOME | Out-Null
                $env:ProgramFiles = '{{Quote(root)}}'
                $payload = {{(fail ? "'exit 23'" : "'[Console]::Write(\"guest-version:\" + $PSVersionTable.PSVersion.Major + \":\" + ''quotes & $ preserved'')'")}}
                $payload = ("# synthetic uploaded file content " + ('x' * 120000) + "`n") + $payload
                $command = ConvertTo-RunnerPowerShell7Command -Script $payload
                & ([scriptblock]::Create($command))
                """);

            Assert.Equal(!fail, result.ExitCode == 0);
            Assert.Contains(fail ? "exit code 23" : "guest-version:7:quotes & $ preserved", result.Output, StringComparison.Ordinal);
            if (fail)
            {
                Assert.DoesNotContain("EXCELMCP_PWSH_COMPLETE", result.Output, StringComparison.Ordinal);
            }
            else
            {
                Assert.Single(result.Output.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries),
                    line => line == "EXCELMCP_PWSH_COMPLETE");
            }
        }
        finally
        {
            var junction = Path.Combine(root, "PowerShell", "7");
            if (Directory.Exists(junction)) { Directory.Delete(junction); }
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void OlapCube_DiscoveryUsesExplicitSyntheticFixtureMarker()
    {
        var setup = File.ReadAllText(Path.Combine(RepoRoot, "scripts", "Start-OlapTestCube.ps1"));
        Assert.Contains("EXCELMCP_TEST_OLAP_SYNTHETIC_FIXTURE = 'true'", setup, StringComparison.Ordinal);
        var discovery = File.ReadAllText(Path.Combine(RepoRoot, "tests", "ExcelMcp.Service.Tests",
            "PersistentServiceConnectionTests.ExternalOlap.cs"));
        Assert.Contains("\"EXCELMCP_TEST_OLAP_SYNTHETIC_FIXTURE\"", discovery, StringComparison.Ordinal);
        Assert.DoesNotContain("cubeName == \"SalesCube\"", discovery, StringComparison.Ordinal);
    }

    [Theory]
    [Trait("RunType", "OnDemand")]
    [InlineData("configure-analytics-oidc.ps1", "-SubscriptionId requested")]
    [InlineData("deploy-appinsights.ps1", "")]
    public async Task AzurePreview_NeverMutatesAndRejectsFailedReads(string script, string arguments)
    {
        var root = NewSandbox();
        try
        {
            var target = Quote(Path.Combine(RepoRoot, "infrastructure", "azure", script));
            var calls = Path.Combine(root, "calls.txt");
            var setup = $$"""
                function az {
                    $call = $args -join ' '
                    Add-Content -LiteralPath '{{Quote(calls)}}' -Value $call -WhatIf:$false
                    $global:LASTEXITCODE = if ($env:MOCK_AZ_FAILURE -eq 'yes') { 23 } else { 0 }
                    switch -Regex ($call) {
                        '^version' { '{"azure-cli":"fixture"}'; break }
                        '^account show' { '{"id":"requested","tenantId":"tenant","name":"fixture"}'; break }
                        '^monitor' { '{"id":"/subscriptions/requested/resourceGroups/fixture/workspaces/test"}'; break }
                        '^ad app list' { '[{"appId":"application"}]'; break }
                        '^ad sp list' { '[{"id":"principal"}]'; break }
                        '^ad app federated-credential list' { '[{"name":"github-main-usage-analytics"}]'; break }
                        '^role assignment list' { '[{"id":"assignment"}]'; break }
                        default { '{}' }
                    }
                }
                function gh { throw 'Preview invoked a GitHub write.' }
                """;
            var preview = await RunAsync(root, setup + $"\n& '{target}' {arguments} -WhatIf");
            Assert.True(preview.ExitCode == 0, preview.Output);
            var recorded = File.ReadAllText(calls);
            Assert.DoesNotContain("account set", recorded, StringComparison.Ordinal);
            Assert.DoesNotContain(" create", recorded, StringComparison.Ordinal);
            Assert.DoesNotContain("login", recorded, StringComparison.Ordinal);
            Assert.DoesNotContain("Configured read-only", preview.Output, StringComparison.Ordinal);
            if (script == "deploy-appinsights.ps1")
            {
                Assert.Contains("deployment sub what-if --subscription requested", recorded, StringComparison.Ordinal);
            }
            File.Delete(calls);
            var failed = await RunAsync(root, setup + $"\n$env:MOCK_AZ_FAILURE='yes'\n& '{target}' {arguments} -WhatIf");
            Assert.NotEqual(0, failed.ExitCode);
            Assert.Single(File.ReadAllLines(calls));
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task AnalyticsOidc_SwitchesToSelectedTenantAndRestoresOriginalSubscription()
    {
        var root = NewSandbox();
        try
        {
            var target = Quote(Path.Combine(
                RepoRoot,
                "infrastructure",
                "azure",
                "configure-analytics-oidc.ps1"));
            var calls = Path.Combine(root, "calls.txt");
            var result = await RunAsync(root, $$"""
                function az {
                    $call = $args -join ' '
                    Add-Content -LiteralPath '{{Quote(calls)}}' -Value $call
                    $global:LASTEXITCODE = 0
                    switch -Regex ($call) {
                        '^account show --only-show-errors$' { '{"id":"original","tenantId":"tenant-a"}'; break }
                        '^account show --subscription selected ' { '{"id":"selected","tenantId":"tenant-b"}'; break }
                        '^monitor' { '{"id":"/subscriptions/selected/resourceGroups/fixture/workspaces/test"}'; break }
                        '^ad app list' { '[{"appId":"application"}]'; break }
                        '^ad sp list' { '[{"id":"principal"}]'; break }
                        '^ad app federated-credential list' { '[{"name":"github-main-usage-analytics"}]'; break }
                        '^role assignment list' { '[{"id":"assignment"}]'; break }
                        default { '{}' }
                    }
                }
                function gh { $global:LASTEXITCODE = 0 }
                & '{{target}}' -SubscriptionId selected -Confirm:$false
                """);
            Assert.True(result.ExitCode == 0, result.Output);
            var recorded = File.ReadAllLines(calls);
            var selectTarget = Array.IndexOf(recorded, "account set --subscription selected");
            var queryApplications = Array.FindIndex(
                recorded,
                call => call.StartsWith("ad app list", StringComparison.Ordinal));
            Assert.True(selectTarget >= 0, "The selected subscription was not activated.");
            Assert.True(queryApplications > selectTarget, "Tenant-scoped commands ran before the selected tenant was active.");
            Assert.Equal("account set --subscription original", recorded[^1]);
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task CaptureEvidence_NeverSavesAndStillQuitsAfterCloseFails()
    {
        var root = NewSandbox();
        try
        {
            var script = Path.Combine(RepoRoot, "videos", "excel-mcp-intro", "Capture-Evidence.ps1");
            var result = await RunAsync(root, $$"""
                $ast = [Management.Automation.Language.Parser]::ParseFile('{{Quote(script)}}', [ref]$null, [ref]$null)
                $outer = $ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] } | Select-Object -Last 1
                $book = [pscustomobject]@{}
                $book | Add-Member ScriptMethod Close { param($save) throw 'simulated close failure' }
                $excel = [pscustomobject]@{}
                $excel | Add-Member ScriptMethod Quit { Write-Output 'quit-after-close-failure' }
                $objects = [Collections.Generic.List[object]]::new()
                & ([scriptblock]::Create('try {} finally ' + $outer.Finally.Extent.Text))
                """);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("quit-after-close-failure", result.Output, StringComparison.Ordinal);
            Assert.DoesNotContain("$book.Save()", File.ReadAllText(script), StringComparison.Ordinal);
        }
        finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TestE2E_DefaultRunSelectsExternalOlapServiceTest(bool failOlapStage)
    {
        var root = NewSandbox();
        try
        {
            var scripts = Directory.CreateDirectory(Path.Combine(root, "scripts")).FullName;
            File.Copy(
                Path.Combine(RepoRoot, "scripts", "Test-E2E.ps1"),
                Path.Combine(scripts, "Test-E2E.ps1"));
            var captured = Path.Combine(root, "stages.jsonl");
            File.WriteAllText(Path.Combine(scripts, "Invoke-TestStage.ps1"), $$"""
                function Invoke-TestStage {
                    param(
                        [string]$Project, [string]$Filter, [string]$ResultsDirectory,
                        [string]$Name, [int]$DeadlineSeconds, [string]$HangTimeout,
                        [hashtable]$Environment, [switch]$ReconcileCases
                    )
                    [pscustomobject]@{
                        project = $Project; filter = $Filter; name = $Name
                        deadline = $DeadlineSeconds; hangTimeout = $HangTimeout
                        reconcileCases = [bool]$ReconcileCases
                        olapConnection = $Environment['EXCELMCP_TEST_OLAP_CONNECTION_STRING']
                        olapCube = $Environment['EXCELMCP_TEST_OLAP_CUBE']
                        olapHierarchy = $Environment['EXCELMCP_TEST_OLAP_HIERARCHY']
                        olapLevel = $Environment['EXCELMCP_TEST_OLAP_LEVEL']
                        syntheticFixture = $Environment['EXCELMCP_TEST_OLAP_SYNTHETIC_FIXTURE']
                    } | ConvertTo-Json -Compress | Add-Content -LiteralPath '{{Quote(captured)}}'
                    if ($Name -eq 'Olap' -and ${{failOlapStage.ToString().ToLowerInvariant()}}) {
                        throw 'Simulated OLAP stage failure.'
                    }
                    $global:LASTEXITCODE = 0
                }
                """);
            File.WriteAllText(
                Path.Combine(scripts, "Stop-ExcelCliService.ps1"),
                "$global:LASTEXITCODE = 0");
            var cubePidFile = Path.Combine(root, "cube.pid");
            var childPidFile = Path.Combine(root, "cube-child.pid");
            File.WriteAllText(Path.Combine(scripts, "Start-OlapTestCube.ps1"), $$"""
                $child = "Start-Process pwsh -NoNewWindow -ArgumentList '-NoProfile','-Command','Start-Sleep 120' -PassThru | ForEach-Object { @{ pid = `$_.Id; started = `$_.StartTime.ToUniversalTime().Ticks } | ConvertTo-Json } | Set-Content '{{Quote(childPidFile)}}'; Start-Sleep 120"
                $process = Start-Process pwsh -NoNewWindow -ArgumentList '-NoProfile', '-Command', $child -PassThru
                @{ pid = $process.Id; started = $process.StartTime.ToUniversalTime().Ticks } | ConvertTo-Json | Set-Content '{{Quote(cubePidFile)}}'
                while (-not (Test-Path '{{Quote(childPidFile)}}')) { Start-Sleep -Milliseconds 100 }
                [pscustomobject]@{
                    Process = $process
                    Settings = @{
                        EXCELMCP_TEST_OLAP_CONNECTION_STRING = 'stub-connection'
                        EXCELMCP_TEST_OLAP_CUBE = 'StubCube'
                        EXCELMCP_TEST_OLAP_HIERARCHY = 'StubHierarchy'
                        EXCELMCP_TEST_OLAP_LEVEL = 'StubLevel'
                        EXCELMCP_TEST_OLAP_SYNTHETIC_FIXTURE = 'true'
                    }
                }
                """);

            var script = Path.Combine(scripts, "Test-E2E.ps1");
            var result = await RunAsync(root, $$"""
                Remove-Item Env:EXCELMCP_TEST_OLAP_CONNECTION_STRING -ErrorAction SilentlyContinue
                Remove-Item Env:EXCELMCP_TEST_OLAP_CUBE -ErrorAction SilentlyContinue
                Remove-Item Env:EXCELMCP_TEST_OLAP_HIERARCHY -ErrorAction SilentlyContinue
                Remove-Item Env:EXCELMCP_TEST_OLAP_LEVEL -ErrorAction SilentlyContinue
                & '{{Quote(script)}}' -SkipBuild -ResultsDirectory '{{Quote(Path.Combine(root, "results"))}}'
                """);

            Assert.Equal(!failOlapStage, result.ExitCode == 0);
            if (failOlapStage) { Assert.Contains("Simulated OLAP stage failure", result.Output, StringComparison.Ordinal); }
            var stages = File.ReadAllLines(captured)
                .Select(line => System.Text.Json.JsonDocument.Parse(line))
                .ToArray();
            try
            {
                Assert.Equal(3, stages.Length);
                var olap = stages[2].RootElement;
                Assert.Equal("Olap", olap.GetProperty("name").GetString());
                Assert.EndsWith(
                    Path.Combine("tests", "ExcelMcp.Service.Tests", "ExcelMcp.Service.Tests.csproj"),
                    olap.GetProperty("project").GetString(),
                    StringComparison.OrdinalIgnoreCase);
                Assert.Equal(
                    "RequiresExcel=true&FullyQualifiedName~PersistentServiceConnectionTests.ExternalOlap",
                    olap.GetProperty("filter").GetString());
                Assert.True(olap.GetProperty("reconcileCases").GetBoolean());
                Assert.Equal(600, olap.GetProperty("deadline").GetInt32());
                Assert.Equal("5m", olap.GetProperty("hangTimeout").GetString());
                Assert.Equal("stub-connection", olap.GetProperty("olapConnection").GetString());
                Assert.Equal("StubCube", olap.GetProperty("olapCube").GetString());
                Assert.Equal("true", olap.GetProperty("syntheticFixture").GetString());
                Assert.Equal(JsonValueKind.Null, stages[0].RootElement.GetProperty("syntheticFixture").ValueKind);
                Assert.Equal(JsonValueKind.Null, stages[0].RootElement.GetProperty("olapConnection").ValueKind);
                foreach (var pidFile in new[] { cubePidFile, childPidFile })
                {
                    using var identity = JsonDocument.Parse(File.ReadAllText(pidFile));
                    var pid = identity.RootElement.GetProperty("pid").GetInt32();
                    Assert.False(IsRunning(pid), $"Cube process {pid} from {Path.GetFileName(pidFile)} is still running.");
                }
            }
            finally
            {
                foreach (var stage in stages) { stage.Dispose(); }
            }
        }
        finally
        {
            try { StopOwnedProcesses(root, "cube.pid", "cube-child.pid"); }
            finally { Directory.Delete(root, true); }
        }
    }

    [Fact]
    public async Task TestE2E_PartialCubeSettingsFailsBeforeStartingReplacementCube()
    {
        var root = NewSandbox();
        try
        {
            var scripts = Directory.CreateDirectory(Path.Combine(root, "scripts")).FullName;
            File.Copy(Path.Combine(RepoRoot, "scripts", "Test-E2E.ps1"), Path.Combine(scripts, "Test-E2E.ps1"));
            File.WriteAllText(Path.Combine(scripts, "Invoke-TestStage.ps1"),
                "function Invoke-TestStage { throw 'Test stage must not run.' }");
            File.WriteAllText(Path.Combine(scripts, "Stop-ExcelCliService.ps1"), "$global:LASTEXITCODE = 0");
            var launched = Path.Combine(root, "launched");
            File.WriteAllText(Path.Combine(scripts, "Start-OlapTestCube.ps1"),
                $"New-Item -ItemType File -Path '{Quote(launched)}'; throw 'Replacement cube started.'");

            var result = await RunAsync(root, """
                $env:EXCELMCP_TEST_OLAP_CONNECTION_STRING = 'configured-connection'
                Remove-Item Env:EXCELMCP_TEST_OLAP_CUBE -ErrorAction SilentlyContinue
                Remove-Item Env:EXCELMCP_TEST_OLAP_HIERARCHY -ErrorAction SilentlyContinue
                Remove-Item Env:EXCELMCP_TEST_OLAP_LEVEL -ErrorAction SilentlyContinue
                & .\scripts\Test-E2E.ps1 -SkipBuild -Stages Olap
                """);

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Configure all four EXCELMCP_TEST_OLAP_", result.Output, StringComparison.Ordinal);
            Assert.False(File.Exists(launched), "Partial configuration must not silently select a different cube.");
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task OlapCube_StartupTimeoutStopsOwnedParentAndChild()
    {
        var root = NewSandbox();
        try
        {
            var parentPid = Path.Combine(root, "parent.pid");
            var childPid = Path.Combine(root, "child.pid");
            var result = await RunAsync(root, $$"""
                . '{{Quote(Path.Combine(RepoRoot, "scripts", "Start-OlapTestCube.ps1"))}}'
                $child = "Start-Process pwsh -NoNewWindow -ArgumentList '-NoProfile','-Command','Start-Sleep 120' -PassThru | ForEach-Object { @{ pid = `$_.Id; started = `$_.StartTime.ToUniversalTime().Ticks } | ConvertTo-Json } | Set-Content '{{Quote(childPid)}}'; Start-Sleep 120"
                $process = Start-Process pwsh -NoNewWindow -ArgumentList '-NoProfile','-Command',$child -PassThru
                @{ pid = $process.Id; started = $process.StartTime.ToUniversalTime().Ticks } | ConvertTo-Json | Set-Content '{{Quote(parentPid)}}'
                $deadline = [DateTime]::UtcNow.AddSeconds(15)
                while (-not (Test-Path '{{Quote(childPid)}}')) {
                    if ([DateTime]::UtcNow -gt $deadline) { $process.Kill($true); throw 'Child startup deadline exceeded.' }
                    Start-Sleep -Milliseconds 100
                }
                $StartupTimeoutSeconds = 1
                $port = 12345
                $logPath = '{{Quote(Path.Combine(root, "cube.log"))}}'
                'startup diagnostic' | Set-Content "$logPath.err"
                try {
                    Wait-OlapCubeStartup -Process $process -LogPath $logPath -Port $port -StartupTimeoutSeconds $StartupTimeoutSeconds
                }
                catch { $_.Exception.ToString() | Write-Output; exit 1 }
                """);

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("did not start on port 12345", result.Output, StringComparison.Ordinal);
            Assert.Contains("startup diagnostic", result.Output, StringComparison.Ordinal);
            foreach (var file in new[] { parentPid, childPid })
            {
                using var identity = JsonDocument.Parse(File.ReadAllText(file));
                Assert.False(IsRunning(identity.RootElement.GetProperty("pid").GetInt32()));
            }
        }
        finally
        {
            try { StopOwnedProcesses(root, "parent.pid", "child.pid"); }
            finally { Directory.Delete(root, true); }
        }
    }

    [Theory]
    [InlineData(23, true)]
    [InlineData(0, false)]
    [InlineData(0, true)]
    public async Task CliWorkflow_ForwardsFocusedAcceptanceAndPropagatesFailure(int exitCode, bool success)
    {
        var root = NewSandbox();
        try
        {
            var script = Path.Combine(RepoRoot, "scripts", "Test-CliWorkflow.ps1");
            File.Copy(script, Path.Combine(root, "Test-CliWorkflow.ps1"));
            var captured = Path.Combine(root, "captured.json");
            File.WriteAllText(Path.Combine(root, "Test-E2E.ps1"), $$"""
                [CmdletBinding()]
                param([switch]$SkipBuild, [string[]]$Stages, [string]$PipeName,
                    [string]$ResultsDirectory, [switch]$KeepCliFiles)
                @{
                    skipBuild = [bool]$SkipBuild; stages = @($Stages); pipe = $PipeName
                    results = $ResultsDirectory; keep = [bool]$KeepCliFiles
                } | ConvertTo-Json | Set-Content -LiteralPath '{{Quote(captured)}}'
                if (-not ${{success.ToString().ToLowerInvariant()}}) { throw 'Simulated acceptance failure.' }
                $global:LASTEXITCODE = {{exitCode}}
                """);
            var result = await RunAsync(root, """
                & .\Test-CliWorkflow.ps1 -PipeName fixture-pipe -ResultsDirectory fixture-results -KeepFile
                """);
            Assert.Equal(success && exitCode == 0, result.ExitCode == 0);
            using var arguments = System.Text.Json.JsonDocument.Parse(File.ReadAllText(captured));
            Assert.True(arguments.RootElement.GetProperty("skipBuild").GetBoolean());
            Assert.True(arguments.RootElement.GetProperty("keep").GetBoolean());
            Assert.Equal("Cli", arguments.RootElement.GetProperty("stages")[0].GetString());
            Assert.Equal(1, arguments.RootElement.GetProperty("stages").GetArrayLength());
            Assert.Equal("fixture-pipe", arguments.RootElement.GetProperty("pipe").GetString());
            Assert.Equal("fixture-results", arguments.RootElement.GetProperty("results").GetString());
        }
        finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task OlapCube_SetupHasDeadlineAndStopsOwnedChildOnTimeout(bool timeout)
    {
        var root = NewSandbox();
        try
        {
            var script = Path.Combine(RepoRoot, "scripts", "Start-OlapTestCube.ps1");
            var childFile = Path.Combine(root, "setup-child.pid");
            var result = await RunAsync(root, $$"""
                . '{{Quote(script)}}'
                $body = @'
                $child = Start-Process pwsh -NoNewWindow -ArgumentList '-NoProfile','-Command','Start-Sleep 120' -PassThru
                @{ pid = $child.Id; started = $child.StartTime.ToUniversalTime().Ticks } | ConvertTo-Json | Set-Content '{{Quote(childFile)}}'
                Start-Sleep 120
                '@
                if (-not ${{timeout.ToString().ToLowerInvariant()}}) { $body = "Write-Output 'setup-success'" }
                Invoke-OlapSetupProcess -Executable 'pwsh' -Arguments @('-NoProfile','-Command',$body) -Deadline ([DateTime]::UtcNow.AddSeconds(5))
                """);

            Assert.Equal(!timeout, result.ExitCode == 0);
            if (timeout)
            {
                Assert.Contains("setup exceeded its hard deadline", result.Output, StringComparison.Ordinal);
                using var identity = JsonDocument.Parse(File.ReadAllText(childFile));
                Assert.False(IsRunning(identity.RootElement.GetProperty("pid").GetInt32()));
            }
            else { Assert.Contains("setup-success", result.Output, StringComparison.Ordinal); }
            var source = File.ReadAllText(script);
            Assert.Contains("-Arguments @('-m', 'venv', $ToolDirectory) -Deadline $setupDeadline", source, StringComparison.Ordinal);
            Assert.Contains("\"atoti==$atotiVersion\") -Deadline $setupDeadline", source, StringComparison.Ordinal);
        }
        finally
        {
            try { StopOwnedProcesses(root, "setup-child.pid"); }
            finally { Directory.Delete(root, true); }
        }
    }

    private static string NewSandbox()
        => Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"ExcelMcp.Automation.{Guid.NewGuid():N}")).FullName;

    private static string Quote(string value) => value.Replace("'", "''", StringComparison.Ordinal);

    private static bool IsRunning(int pid)
    {
        try
        {
            using var process = Process.GetProcessById(pid);
            return !process.HasExited;
        }
        catch (ArgumentException) { return false; }
    }

    private static void StopOwnedProcesses(string root, params string[] identityFiles)
    {
        foreach (var file in identityFiles)
        {
            var path = Path.Combine(root, file);
            if (!File.Exists(path)) { continue; }
            using var identity = JsonDocument.Parse(File.ReadAllText(path));
            Process process;
            try { process = Process.GetProcessById(identity.RootElement.GetProperty("pid").GetInt32()); }
            catch (ArgumentException) { continue; }
            using (process)
            {
                _ = process.Handle;
                if (process.HasExited || process.StartTime.ToUniversalTime().Ticks !=
                    identity.RootElement.GetProperty("started").GetInt64()) { continue; }
                process.Kill(entireProcessTree: true);
                Assert.True(process.WaitForExit(10000), $"Owned process {process.Id} did not stop.");
            }
        }
    }

    private static async Task<(int ExitCode, string Output)> RunAsync(string root, string body)
    {
        var script = Path.Combine(root, $"test-{Guid.NewGuid():N}.ps1");
        File.WriteAllText(script, "$ErrorActionPreference='Stop'\n$global:LASTEXITCODE=0\n" + body + "\nexit $LASTEXITCODE");
        var result = await new Sbroenne.ExcelMcp.Build.ProcessRunner(root).RunAsync(
            "pwsh", ["-NoProfile", "-File", script], TimeSpan.FromSeconds(60),
            new Dictionary<string, string>
            {
                ["EXCELMCP_BUILD_ROOT"] = RepoRoot,
                ["EXCELMCP_BUILD_DLL"] = typeof(Sbroenne.ExcelMcp.Build.ValidationPolicy).Assembly.Location
            });
        return (result.ExitCode, result.Output + result.Error);
    }

    private static string FindRepoRoot()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory != null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln"))) { return directory.FullName; }
        }
        throw new DirectoryNotFoundException("Repository root was not found.");
    }
}
