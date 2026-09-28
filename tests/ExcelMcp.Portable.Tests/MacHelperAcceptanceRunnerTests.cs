using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacHelperAcceptanceRunnerTests
{
    [Fact]
    public void PublicVbaRunner_UsesOnlyGuardedCliAndMcpAcceptance()
    {
        var script = File.ReadAllText(
            Path.Combine(
                FindRepository(),
                "scripts",
                "Test-MacVbaPublicAcceptance.ps1"));

        Assert.Contains("EXCELMCP_MAC_VBA_CANDIDATE_ACTIONS", script, StringComparison.Ordinal);
        Assert.Contains("'vba.list,vba.view,vba.import,vba.update,vba.delete,vba.run'", script, StringComparison.Ordinal);
        Assert.Contains("method = 'tools/call'", script, StringComparison.Ordinal);
        Assert.Contains("procedure_name = $MarkerProcedure", script, StringComparison.Ordinal);
        Assert.Contains("action = 'get-values'", script, StringComparison.Ordinal);
        Assert.Contains("save = $false", script, StringComparison.Ordinal);
        Assert.Contains("publicCommandAcceptance = $true", script, StringComparison.Ordinal);
        Assert.DoesNotContain("vbaProject.bin", script, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ValidateOnly_RequiresExplicitSecurityAndProvenanceConfirmations()
    {
        using var fixture = new RunnerFixture();

        var result = await fixture.RunAsync();

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains(
            "MacroApprovalConfirmed",
            result.StandardError + result.StandardOutput,
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task ValidateOnly_RejectsImpreciseHelperAndWorkbookInputs()
    {
        using var fixture = new RunnerFixture(
            helperName: "Imposter.xlam",
            workbookName: "fixture.xlsx");

        var result = await fixture.RunAsync(
            "-MacroApprovalConfirmed",
            "-VbaProjectTrustConfirmed",
            "-ExcelAuthoredWorkbookConfirmed");

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains(
            "ExcelMcpHelper.xlam",
            result.StandardError + result.StandardOutput,
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task ValidateOnly_EmitsNonProofPlanWithExactCommittedSourceAndBackends()
    {
        using var fixture = new RunnerFixture();

        var result = await fixture.RunAsync(
            "-MacroApprovalConfirmed",
            "-VbaProjectTrustConfirmed",
            "-ExcelAuthoredWorkbookConfirmed");

        Assert.Equal(0, result.ExitCode);
        using var output = JsonDocument.Parse(result.StandardOutput);
        var root = output.RootElement;
        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("validation-only", root.GetProperty("status").GetString());
        Assert.Equal("direct-helper-engine", root.GetProperty("acceptanceScope").GetString());
        Assert.False(root.GetProperty("runtimeProof").GetBoolean());
        Assert.False(root.GetProperty("publicCommandAcceptance").GetBoolean());
        Assert.Equal(
            Path.GetFullPath(fixture.HelperPath),
            root.GetProperty("helperPath").GetString());
        Assert.Equal(
            Path.GetFullPath(fixture.WorkbookPath),
            root.GetProperty("workbookPath").GetString());
        Assert.EndsWith(
            "src/ExcelMcp.Service/Mac/ExcelMcpHelper.bas",
            root.GetProperty("helperSourcePath").GetString(),
            StringComparison.Ordinal);
        Assert.Equal(
            ["cli", "mcp"],
            root.GetProperty("entryPoints").EnumerateArray()
                .Select(item => item.GetString()!).ToArray());
        Assert.Equal(
            ["helper.capabilities", "helper.inspect-engines", "powerquery", "vba"],
            root.GetProperty("phases").EnumerateArray()
                .Select(item => item.GetString()!).ToArray());
    }

    [Fact]
    public void Runner_IsOptInAndKeepsDirectAcceptanceSeparateFromPublicParity()
    {
        var script = File.ReadAllText(Path.Combine(
            FindRepository(),
            "scripts",
            "Test-MacHelperAcceptance.ps1"));

        Assert.Contains("EXCELMCP_MAC_VBA_HELPER_PATH", script, StringComparison.Ordinal);
        Assert.Contains("ExcelMcpHelper.bas", script, StringComparison.Ordinal);
        Assert.Contains("Sbroenne.ExcelMcp.McpServer.dll", script, StringComparison.Ordinal);
        Assert.Contains("--excelmcp-mac-automation", script, StringComparison.Ordinal);
        Assert.Contains("helper.dispatch", script, StringComparison.Ordinal);
        Assert.Contains("helper version does not match protocol 1 / helper 1.3.0", script, StringComparison.Ordinal);
        Assert.Contains("helper.inspect-engines", script, StringComparison.Ordinal);
        Assert.Contains("Assert-EngineInspection", script, StringComparison.Ordinal);
        Assert.DoesNotContain("helper 1.1.0", script, StringComparison.Ordinal);
        Assert.Contains("destination = 'connection-only'", script, StringComparison.Ordinal);
        Assert.Contains("refresh = $false", script, StringComparison.Ordinal);
        Assert.Contains("publicCommandAcceptance = $false", script, StringComparison.Ordinal);
        Assert.Contains("Invalidate-RunnerSession", script, StringComparison.Ordinal);
        Assert.Contains("RECOVERY_REQUIRED", script, StringComparison.Ordinal);
        Assert.Contains("rollback_failed", script, StringComparison.Ordinal);
        Assert.Contains("$manualReconciliationRequired = $true", script, StringComparison.Ordinal);
        Assert.Contains("$queryCreated = $false", script, StringComparison.Ordinal);
        Assert.Contains("if ($script:queryCreated)", script, StringComparison.Ordinal);
        Assert.Contains("if ($script:moduleCreated)", script, StringComparison.Ordinal);
        Assert.DoesNotContain("-IncludePowerQueryFixtures", script, StringComparison.Ordinal);
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
            $"excelmcp-helper-runner-{Guid.NewGuid():N}");

        public RunnerFixture(
            string helperName = "ExcelMcpHelper.xlam",
            string workbookName = "fixture.xlsm")
        {
            Directory.CreateDirectory(_directory);
            HelperPath = Path.Combine(_directory, helperName);
            WorkbookPath = Path.Combine(_directory, workbookName);
            File.WriteAllBytes(HelperPath, []);
            File.WriteAllBytes(WorkbookPath, []);
        }

        public string HelperPath { get; }
        public string WorkbookPath { get; }

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
                "Test-MacHelperAcceptance.ps1"));
            start.ArgumentList.Add("-ValidateOnly");
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
