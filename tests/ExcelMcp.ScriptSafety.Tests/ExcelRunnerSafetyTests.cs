using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "AutomationSafety")]
public sealed class ExcelRunnerSafetyTests
{
    public static IEnumerable<object[]> Scripts()
    {
        var directory = Path.Combine(FindRoot(), "scripts", "tests");
        var scripts = Directory.GetFiles(directory, "*.tests.ps1").Order(StringComparer.Ordinal).ToArray();
        if (scripts.Length == 0) { throw new InvalidOperationException("Runner safety scripts were not discovered."); }
        foreach (var script in scripts)
        {
            yield return [script, "pwsh"];
            if (OperatingSystem.IsWindows()) { yield return [script, "powershell"]; }
        }
    }

    [Theory]
    [MemberData(nameof(Scripts))]
    public async Task OfflineRunnerSafety_Passes(string script, string shell)
    {
        var info = new ProcessStartInfo(shell)
        {
            WorkingDirectory = FindRoot(),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        if (shell == "powershell") { info.Environment.Remove("PSModulePath"); }
        foreach (var argument in new[] { "-NoProfile", "-NonInteractive", "-File", script }) { info.ArgumentList.Add(argument); }
        using var process = Process.Start(info);
        Assert.NotNull(process);
        var output = process.StandardOutput.ReadToEndAsync();
        var errors = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
        try { await process.WaitForExitAsync(timeout.Token); }
        catch (OperationCanceledException) when (timeout.IsCancellationRequested)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw new TimeoutException("Offline runner safety test exceeded its deadline.");
        }
        Assert.True(process.ExitCode == 0, await output + await errors);
    }

    private static string FindRoot()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln"))) { return directory.FullName; }
        }
        throw new DirectoryNotFoundException("Repository root was not found.");
    }
}
