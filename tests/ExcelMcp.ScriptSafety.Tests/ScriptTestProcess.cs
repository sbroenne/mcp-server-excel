using System.Diagnostics;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

internal sealed class ScriptTestProcess(string root)
{
    internal sealed record Result(int ExitCode, string Output, string Error);

    internal async Task<Result> RunAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
        IReadOnlyDictionary<string, string>? environment = null)
    {
        var info = new ProcessStartInfo(executable)
        {
            WorkingDirectory = root,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        foreach (var argument in arguments) { info.ArgumentList.Add(argument); }
        foreach (var key in info.Environment.Keys.Where(key => key.StartsWith("GIT_", StringComparison.OrdinalIgnoreCase)).ToArray())
        {
            info.Environment.Remove(key);
        }
        if (environment is not null)
        {
            foreach (var (key, value) in environment) { info.Environment[key] = value; }
        }
        using var process = Process.Start(info) ?? throw new InvalidOperationException($"Cannot start {executable}.");
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        try
        {
            await Task.WhenAll(process.WaitForExitAsync(), output, error).WaitAsync(deadline);
        }
        catch (TimeoutException)
        {
            if (!process.HasExited) { process.Kill(entireProcessTree: true); }
            await Task.WhenAll(process.WaitForExitAsync(), output, error).WaitAsync(TimeSpan.FromSeconds(10));
            throw;
        }
        return new Result(process.ExitCode, await output, await error);
    }
}
