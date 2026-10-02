using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacOfficeJsAcceptanceRunnerTests
{
    [Fact]
    public async Task ValidateOnly_RequiresExplicitUserManagedSetupAndDesktopSlot()
    {
        using var fixture = new RunnerFixture();

        var result = await fixture.RunAsync();

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains(
            "UserSetupConfirmed",
            result.StandardError + result.StandardOutput,
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task ValidateOnly_RejectsMissingCandidateAllowlistEntries()
    {
        using var fixture = new RunnerFixture(["bridge.health"]);

        var result = await fixture.RunAsync(
            "-UserSetupConfirmed",
            "-CandidateAllowlistConfirmed",
            "-ExcelSlotConfirmed");

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains(
            "pivottablefield.remove-field",
            result.StandardError + result.StandardOutput,
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task ValidateOnly_EmitsNonProofPublicEntryPointPlan()
    {
        using var fixture = new RunnerFixture(RunnerFixture.RequiredActions);

        var result = await fixture.RunAsync(
            "-UserSetupConfirmed",
            "-CandidateAllowlistConfirmed",
            "-ExcelSlotConfirmed");

        Assert.Equal(0, result.ExitCode);
        using var output = JsonDocument.Parse(result.StandardOutput);
        var root = output.RootElement;
        Assert.Equal("validation-only", root.GetProperty("status").GetString());
        Assert.Equal(
            "public-officejs-pivottable-candidates",
            root.GetProperty("acceptanceScope").GetString());
        Assert.False(root.GetProperty("runtimeProof").GetBoolean());
        Assert.False(root.GetProperty("publicCommandAcceptance").GetBoolean());
        Assert.Equal(
            ["cli", "mcp"],
            root.GetProperty("entryPoints").EnumerateArray()
                .Select(item => item.GetString()!).ToArray());
        Assert.Equal(
            RunnerFixture.RequiredActions,
            root.GetProperty("requiredActions").EnumerateArray()
                .Select(item => item.GetString()!).ToArray());
    }

    [Fact]
    public void Runner_DoesNotAutomateTrustSetupOrEnableCandidates()
    {
        var script = File.ReadAllText(Path.Combine(
            FindRepository(),
            "scripts",
            "Test-MacOfficeJsAcceptance.ps1"));

        Assert.Contains("UserSetupConfirmed", script, StringComparison.Ordinal);
        Assert.Contains("CandidateAllowlistConfirmed", script, StringComparison.Ordinal);
        Assert.Contains("ExcelSlotConfirmed", script, StringComparison.Ordinal);
        Assert.Contains("Sbroenne.ExcelMcp.McpServer.dll", script, StringComparison.Ordinal);
        Assert.Contains("publicCommandAcceptance = $true", script, StringComparison.Ordinal);
        Assert.DoesNotContain("security add-trusted-cert", script, StringComparison.Ordinal);
        Assert.DoesNotContain("certutil", script, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("WriteAllText($configPath", script, StringComparison.Ordinal);
        Assert.DoesNotContain("open -a", script, StringComparison.Ordinal);
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
            $"excelmcp-officejs-runner-{Guid.NewGuid():N}");

        public static readonly string[] RequiredActions =
        [
            "pivottablefield.remove-field",
            "pivottablefield.set-field-name",
            "pivottablefield.set-field-format",
            "pivottablefield.set-field-filter",
            "pivottablefield.sort-field",
            "pivottablecalc.get-data",
            "pivottablecalc.set-layout",
            "pivottablecalc.set-subtotals",
            "pivottablecalc.set-grand-totals"
        ];

        public RunnerFixture(string[]? enabledActions = null)
        {
            Directory.CreateDirectory(_directory);
            WorkbookPath = Path.Combine(_directory, "acceptance.xlsx");
            ConfigPath = Path.Combine(_directory, "bridge.json");
            File.WriteAllBytes(WorkbookPath, []);
            File.WriteAllText(
                ConfigPath,
                JsonSerializer.Serialize(new
                {
                    enabledActions = enabledActions ?? RequiredActions
                }));
        }

        public string WorkbookPath { get; }
        public string ConfigPath { get; }

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
                "Test-MacOfficeJsAcceptance.ps1"));
            start.ArgumentList.Add("-ValidateOnly");
            start.ArgumentList.Add("-WorkbookPath");
            start.ArgumentList.Add(WorkbookPath);
            start.ArgumentList.Add("-BridgeConfigPath");
            start.ArgumentList.Add(ConfigPath);
            start.ArgumentList.Add("-CliPivotTable");
            start.ArgumentList.Add("CliPivot");
            start.ArgumentList.Add("-CliRemovalPivotTable");
            start.ArgumentList.Add("CliRemovalPivot");
            start.ArgumentList.Add("-McpPivotTable");
            start.ArgumentList.Add("McpPivot");
            start.ArgumentList.Add("-McpRemovalPivotTable");
            start.ArgumentList.Add("McpRemovalPivot");
            start.ArgumentList.Add("-RowField");
            start.ArgumentList.Add("Region");
            start.ArgumentList.Add("-ValueField");
            start.ArgumentList.Add("Sales");
            start.ArgumentList.Add("-SelectedItem");
            start.ArgumentList.Add("North");
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
