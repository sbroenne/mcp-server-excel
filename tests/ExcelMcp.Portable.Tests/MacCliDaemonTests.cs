using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacCliFactAttribute : FactAttribute
{
    public MacCliFactAttribute()
    {
        if (!OperatingSystem.IsMacOS())
        {
            Skip = "Requires the macOS CLI build, but does not open Excel workbooks.";
        }
    }
}

public sealed class MacCliDaemonTests
{
    [MacCliFact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "MacDaemon")]
    public async Task FrameworkDependentCli_StartsAndStopsItsPrivateDaemon()
    {
        var root = MacExcelE2ETests.FindRepository();
        var assembly = Path.Combine(root, "src", "ExcelMcp.CLI", "bin", "Release", "net10.0", "excelcli.dll");
        Assert.True(File.Exists(assembly), "Build the Release macOS CLI before running its daemon regression.");
        var pipe = $"em-{Guid.NewGuid():N}";
        try
        {
            await InvokeAsync(assembly, pipe, "start");
            var running = await InvokeAsync(assembly, pipe, "status");
            Assert.True(running.GetProperty("running").GetBoolean());
            Assert.Equal(0, running.GetProperty("sessionCount").GetInt32());
            Assert.True(running.GetProperty("processId").GetInt32() > 0);

            await InvokeAsync(assembly, pipe, "stop");
            var stopped = await InvokeAsync(assembly, pipe, "status");
            Assert.False(stopped.GetProperty("running").GetBoolean());
            Assert.Equal(0, stopped.GetProperty("processId").GetInt32());
        }
        finally
        {
            await InvokeAsync(assembly, pipe, "stop");
        }
    }

    private static async Task<JsonElement> InvokeAsync(string assembly, string pipe, string action)
    {
        var start = new ProcessStartInfo("dotnet")
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        start.Environment["EXCELMCP_CLI_PIPE"] = pipe;
        foreach (var argument in new[] { assembly, "-q", "service", action })
        {
            start.ArgumentList.Add(argument);
        }
        using var process = Process.Start(start)
            ?? throw new InvalidOperationException("CLI process did not start.");
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try
        {
            await process.WaitForExitAsync(deadline.Token);
        }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw;
        }
        var output = await stdout;
        Assert.True(process.ExitCode == 0, $"{action}: {output} {await stderr}");
        using var document = JsonDocument.Parse(output);
        var result = document.RootElement.Clone();
        Assert.True(result.GetProperty("success").GetBoolean(), output);
        return result;
    }
}
