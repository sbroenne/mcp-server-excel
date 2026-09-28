using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacPowerQueryPublicAcceptanceRunnerTests
{
    [Fact]
    public async Task ValidateOnly_RequiresExplicitSetupTrustAndSlotConfirmations()
    {
        using var fixture = new RunnerFixture();

        var result = await fixture.RunAsync();

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains(
            "HelperInstalledTrustedConfirmed",
            result.StandardError + result.StandardOutput,
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task ValidateOnly_RejectsImpreciseHelperAndWorkbookInputs()
    {
        using var fixture = new RunnerFixture(
            helperName: "Imposter.xlam",
            workbookName: "fixture.csv");

        var result = await fixture.RunAsync(RunnerFixture.Confirmations);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains(
            "ExcelMcpHelper.xlam",
            result.StandardError + result.StandardOutput,
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task ValidateOnly_EmitsNonProofPublicLifecyclePlan()
    {
        using var fixture = new RunnerFixture();

        var result = await fixture.RunAsync(RunnerFixture.Confirmations);

        Assert.Equal(0, result.ExitCode);
        using var output = JsonDocument.Parse(result.StandardOutput);
        var root = output.RootElement;
        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("validation-only", root.GetProperty("status").GetString());
        Assert.Equal(
            "public-powerquery-lifecycle-cli-mcp",
            root.GetProperty("acceptanceScope").GetString());
        Assert.False(root.GetProperty("runtimeProof").GetBoolean());
        Assert.False(root.GetProperty("publicCommandAcceptance").GetBoolean());
        Assert.Equal("1.4.0", root.GetProperty("helperVersion").GetString());
        Assert.Equal(
            ["cli", "mcp"],
            root.GetProperty("entryPoints").EnumerateArray()
                .Select(item => item.GetString()!).ToArray());
        Assert.Equal(
            RunnerFixture.RequiredActions,
            root.GetProperty("requiredActions").EnumerateArray()
                .Select(item => item.GetString()!).ToArray());
        Assert.Equal(
            ["load-to-data-model", "load-to-both"],
            root.GetProperty("requiredUnsupportedDestinations").EnumerateArray()
                .Select(item => item.GetString()!).ToArray());
    }

    [Fact]
    public void Runner_UsesGuardedPublicCliAndMcpContractsOnly()
    {
        var script = File.ReadAllText(Path.Combine(
            FindRepository(),
            "scripts",
            "Test-MacPowerQueryPublicAcceptance.ps1"));

        Assert.Contains("EXCELMCP_MAC_POWERQUERY_CANDIDATE_ACTIONS", script, StringComparison.Ordinal);
        Assert.Contains(
            "$candidateActions = $requiredActions -join ','",
            script,
            StringComparison.Ordinal);
        Assert.All(
            RunnerFixture.RequiredActions,
            action => Assert.Contains($"'{action}'", script, StringComparison.Ordinal));
        Assert.Contains("method = 'tools/call'", script, StringComparison.Ordinal);
        Assert.Contains("'powerquery'", script, StringComparison.Ordinal);
        Assert.Contains("load-to-data-model", script, StringComparison.Ordinal);
        Assert.Contains("load-to-both", script, StringComparison.Ordinal);
        Assert.Contains("publicCommandAcceptance = $true", script, StringComparison.Ordinal);
        Assert.Contains("save = $false", script, StringComparison.Ordinal);
        Assert.Contains("Assert-DedicatedWorkbookIsEmpty", script, StringComparison.Ordinal);
        Assert.Contains("TIMEOUT_UNCERTAIN", script, StringComparison.Ordinal);
        Assert.Contains("RECOVERY_REQUIRED", script, StringComparison.Ordinal);
        Assert.DoesNotContain("--excelmcp-mac-automation", script, StringComparison.Ordinal);
        Assert.DoesNotContain("helper.dispatch", script, StringComparison.Ordinal);
        Assert.DoesNotContain("provenMethods", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Test-MacE2E.ps1", script, StringComparison.Ordinal);
    }

    [Fact]
    public void Runner_VerifiesLoadedValuesAcrossRefreshAndSavedReopen()
    {
        var script = ReadRunner();

        Assert.Contains("function Assert-LoadedValues", script, StringComparison.Ordinal);
        Assert.Contains("'range', 'get-values'", script, StringComparison.Ordinal);
        Assert.Contains("action = 'get-values'", script, StringComparison.Ordinal);
        Assert.Contains("'A1:A2'", script, StringComparison.Ordinal);
        Assert.Contains("Assert-LoadedValues $loaded 'original'", script, StringComparison.Ordinal);
        Assert.Contains("Assert-LoadedValues $loaded 'updated'", script, StringComparison.Ordinal);
        Assert.Contains("CLI refreshed checkpoint", script, StringComparison.Ordinal);
        Assert.Contains("CLI refresh-all checkpoint", script, StringComparison.Ordinal);
        Assert.Contains("MCP refreshed checkpoint", script, StringComparison.Ordinal);
        Assert.Contains("MCP refresh-all checkpoint", script, StringComparison.Ordinal);
    }

    [Fact]
    public void Runner_BoundsMcpProtocolAndDrainsStandardError()
    {
        var script = ReadRunner();

        Assert.Contains("StandardError.ReadToEndAsync()", script, StringComparison.Ordinal);
        Assert.Contains("[Diagnostics.Stopwatch]::StartNew()", script, StringComparison.Ordinal);
        Assert.Contains("$remaining", script, StringComparison.Ordinal);
        Assert.DoesNotContain(
            "[TimeSpan]::FromSeconds($OperationTimeoutSeconds)).GetAwaiter().GetResult()",
            script,
            StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(0, false)]
    public async Task Runner_ShutsDownOnlyItsPrivateCliPipe(int exitCode, bool success)
    {
        var script = ReadRunner();

        Assert.DoesNotContain("Stop-Process", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Get-Process", script, StringComparison.Ordinal);
        Assert.Contains("$start.Environment[$entry.Key] = $entry.Value", script, StringComparison.Ordinal);
        using var fixture = new RunnerFixture();
        var result = await fixture.RunCleanupAsync(exitCode, success);
        Assert.Equal(0, result.ExitCode);
        using var document = JsonDocument.Parse(result.StandardOutput);
        var receipt = document.RootElement;
        Assert.Equal(1, receipt.GetProperty("calls").GetInt32());
        Assert.Equal(
            ["fixture-cli.dll", "-q", "service", "stop"],
            receipt.GetProperty("arguments").EnumerateArray().Select(item => Assert.IsType<string>(item.GetString())).ToArray());
        var pipe = Assert.IsType<string>(receipt.GetProperty("pipe").GetString());
        var macSocketPath = "/var/folders/aa/" + new string('a', 30) + "/T/CoreFxPipe_" + pipe;
        Assert.InRange(System.Text.Encoding.UTF8.GetByteCount(macSocketPath), 1, 104);
        Assert.Equal(exitCode != 0 || !success, receipt.GetProperty("failed").GetBoolean());
        if (exitCode != 0 || !success)
        {
            Assert.Contains("Private CLI daemon cleanup", receipt.GetProperty("error").GetString(), StringComparison.Ordinal);
        }
    }

    [Fact]
    public void Runner_PreservesWorkingCopyUnlessExactCloseIsConfirmed()
    {
        var script = ReadRunner();

        Assert.Contains("$exactlyClosed = $true", script, StringComparison.Ordinal);
        Assert.Contains("$ExactlyClosed.Value = $false", script, StringComparison.Ordinal);
        Assert.Contains("$ExactlyClosed.Value = $true", script, StringComparison.Ordinal);
        Assert.DoesNotContain(
            "if ($null -eq $session -and -not $script:uncertain)",
            script,
            StringComparison.Ordinal);
        Assert.Contains(
            "could not confirm exact close. Preserve",
            script,
            StringComparison.Ordinal);
    }

    private static string ReadRunner()
    {
        return File.ReadAllText(Path.Combine(
            FindRepository(),
            "scripts",
            "Test-MacPowerQueryPublicAcceptance.ps1"));
    }

    private static string FindRepository()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null
               && !File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            directory = directory.Parent;
        }

        return directory?.FullName
            ?? throw new InvalidOperationException("Repository root not found.");
    }

    private sealed class RunnerFixture : IDisposable
    {
        private readonly string _directory = Path.Combine(
            Path.GetTempPath(),
            $"excelmcp-pq-public-runner-{Guid.NewGuid():N}");

        public static readonly string[] Confirmations =
        [
            "-HelperInstalledTrustedConfirmed",
            "-ExcelAuthoredWorkbookConfirmed",
            "-DedicatedWorkbookConfirmed",
            "-ExcelSlotConfirmed"
        ];

        public static readonly string[] RequiredActions =
        [
            "powerquery.create",
            "powerquery.update",
            "powerquery.rename",
            "powerquery.delete",
            "powerquery.refresh",
            "powerquery.refresh-all",
            "powerquery.load-to",
            "powerquery.unload",
            "powerquery.evaluate"
        ];

        public RunnerFixture(
            string helperName = "ExcelMcpHelper.xlam",
            string workbookName = "fixture.xlsx")
        {
            Directory.CreateDirectory(_directory);
            HelperPath = Path.Combine(_directory, helperName);
            WorkbookPath = Path.Combine(_directory, workbookName);
            File.WriteAllBytes(HelperPath, []);
            File.WriteAllBytes(WorkbookPath, []);
        }

        public string HelperPath { get; }
        public string WorkbookPath { get; }

        public async Task<ProcessResult> RunCleanupAsync(int exitCode, bool success)
        {
            var start = new ProcessStartInfo("pwsh")
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = _directory
            };
            start.Environment["RUNNER_SCRIPT_PATH"] = Path.Combine(
                FindRepository(), "scripts", "Test-MacPowerQueryPublicAcceptance.ps1");
            start.Environment["FIXTURE_EXIT_CODE"] = exitCode.ToString(System.Globalization.CultureInfo.InvariantCulture);
            start.Environment["FIXTURE_SUCCESS"] = success ? "true" : "false";
            start.ArgumentList.Add("-NoProfile");
            start.ArgumentList.Add("-Command");
            start.ArgumentList.Add("""
                $ErrorActionPreference = 'Stop'
                $names = @('Stop-PrivateCliDaemon', 'Invoke-Cli', 'Invoke-CliRaw', 'ConvertFrom-StrictJson')
                $ast = [Management.Automation.Language.Parser]::ParseFile($env:RUNNER_SCRIPT_PATH, [ref]$null, [ref]$null)
                $functions = $ast.FindAll({
                    param($node)
                    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name
                }, $true)
                foreach ($function in $functions) { . ([scriptblock]::Create($function.Extent.Text)) }
                $helperPath = 'fixture-helper.xlam'
                $candidateActions = 'powerquery.create'
                $environmentAssignment = $ast.Find({
                    param($node)
                    $node -is [Management.Automation.Language.AssignmentStatementAst] -and
                    $node.Left -is [Management.Automation.Language.VariableExpressionAst] -and
                    $node.Left.VariablePath.UserPath -ceq 'environment'
                }, $true)
                . ([scriptblock]::Create($environmentAssignment.Extent.Text))
                $cliAssembly = 'fixture-cli.dll'
                $script:calls = 0
                $script:observedArguments = @()
                function Join-Path { throw 'Windows-only cleanup must not be invoked.' }
                function Invoke-BoundedProcess {
                    param([string]$FileName, [string[]]$Arguments)
                    if ($FileName -cne 'dotnet') { throw 'Unexpected process.' }
                    $script:calls++
                    $script:observedArguments = $Arguments
                    return @{
                        exitCode = [int]$env:FIXTURE_EXIT_CODE
                        stdout = '{"success":' + $env:FIXTURE_SUCCESS + '}'
                        stderr = 'fixture stderr'
                    }
                }
                $failed = $false
                $message = ''
                try { Stop-PrivateCliDaemon } catch { $failed = $true; $message = $_.Exception.Message }
                @{
                    calls = $script:calls
                    arguments = $script:observedArguments
                    pipe = $environment.EXCELMCP_CLI_PIPE
                    failed = $failed
                    error = $message
                } | ConvertTo-Json -Compress
                """);
            using var process = Process.Start(start)
                ?? throw new InvalidOperationException("PowerShell did not start.");
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            try
            {
                await process.WaitForExitAsync(timeout.Token);
            }
            catch (OperationCanceledException)
            {
                process.Kill(entireProcessTree: true);
                await process.WaitForExitAsync();
                throw;
            }
            return new ProcessResult(process.ExitCode, await stdout, await stderr);
        }

        public async Task<ProcessResult> RunAsync(params string[] confirmations)
        {
            var start = new ProcessStartInfo("pwsh")
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = FindRepository()
            };
            start.ArgumentList.Add("-NoProfile");
            start.ArgumentList.Add("-File");
            start.ArgumentList.Add(Path.Combine(
                FindRepository(),
                "scripts",
                "Test-MacPowerQueryPublicAcceptance.ps1"));
            start.ArgumentList.Add("-ValidateOnly");
            start.ArgumentList.Add("-SkipBuild");
            start.ArgumentList.Add("-HelperPath");
            start.ArgumentList.Add(HelperPath);
            start.ArgumentList.Add("-WorkbookPath");
            start.ArgumentList.Add(WorkbookPath);
            foreach (var confirmation in confirmations)
            {
                start.ArgumentList.Add(confirmation);
            }

            using var process = Process.Start(start)
                ?? throw new InvalidOperationException("PowerShell did not start.");
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(30));
            return new ProcessResult(
                process.ExitCode,
                await stdout,
                await stderr);
        }

        public void Dispose()
        {
            Directory.Delete(_directory, recursive: true);
        }
    }

    private sealed record ProcessResult(
        int ExitCode,
        string StandardOutput,
        string StandardError);
}
