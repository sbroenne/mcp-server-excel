using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "AutomationSafety")]
public sealed class IssueTriageScriptTests
{
    [Fact]
    public async Task TriageAndCodingAgentHandoff_RespectRestrictedWorkflowContract()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !File.Exists(Path.Combine(root.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            root = root.Parent;
        }
        Assert.NotNull(root);
        var start = new ProcessStartInfo("node")
        {
            WorkingDirectory = root.FullName,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        start.ArgumentList.Add("--test");
        start.ArgumentList.Add(Path.Combine("tests", "ExcelMcp.ScriptSafety.Tests", "IssueTriage.test.mjs"));
        using var process = Process.Start(start);
        Assert.NotNull(process);
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        try
        {
            await process.WaitForExitAsync(deadline.Token);
        }
        catch (OperationCanceledException) when (deadline.IsCancellationRequested)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw new TimeoutException($"Issue triage regressions exceeded their deadline.\n{await stdout}\n{await stderr}");
        }
        Assert.True(process.ExitCode == 0,
            $"Issue triage regressions exited {process.ExitCode}.\n{await stdout}\n{await stderr}");
    }
}
