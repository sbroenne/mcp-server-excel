using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "PluginPublication")]
public sealed class PluginPublicationScriptTests
{
    [Fact]
    public async Task PublicationAndMarketplaceScripts_PassDisposableNoWriteRegressions()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !File.Exists(Path.Combine(root.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            root = root.Parent;
        }
        Assert.NotNull(root);
        var info = new ProcessStartInfo("node")
        {
            WorkingDirectory = root.FullName,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        info.ArgumentList.Add("--test");
        info.ArgumentList.Add("tests/ExcelMcp.SkillGeneration.Tests/PluginPublication.test.mjs");
        using var process = Process.Start(info);
        Assert.NotNull(process);
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(3));
        try
        {
            await process.WaitForExitAsync(deadline.Token);
        }
        catch (OperationCanceledException)
        {
            process.Kill(true);
            await process.WaitForExitAsync();
            throw new TimeoutException("Plugin publication script tests exceeded three minutes.");
        }
        Assert.True(process.ExitCode == 0, await stdout + Environment.NewLine + await stderr);
    }
}
