using System.Diagnostics;

namespace Sbroenne.ExcelMcp.Build;

public sealed record ProcessResult(int ExitCode, string Output, string Error);

public interface IProcessRunner
{
    Task<ProcessResult> RunAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
        IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false);
    Task<ProcessResult> CheckedAsync(string executable, IEnumerable<string> arguments, TimeSpan deadline,
        IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false);
}

public sealed class ProcessRunner(string root) : IProcessRunner
{
    private readonly Action<Process> _terminate = static process => process.Kill(entireProcessTree: true);

    internal ProcessRunner(string root, Action<Process> terminate) : this(root)
    {
        ArgumentNullException.ThrowIfNull(terminate);
        _terminate = terminate;
    }

    public async Task<ProcessResult> RunAsync(
        string executable, IEnumerable<string> arguments, TimeSpan deadline,
        IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false)
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
        if (!preserveGitContext)
        {
            foreach (var key in info.Environment.Keys.Where(key => key.StartsWith("GIT_", StringComparison.OrdinalIgnoreCase)).ToArray())
            {
                info.Environment.Remove(key);
            }
        }
        if (environment is not null)
        {
            foreach (var (key, value) in environment) { info.Environment[key] = value; }
        }
        using var process = Process.Start(info) ?? throw new InvalidOperationException($"Could not start {executable}.");
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(deadline);
        try
        {
            await process.WaitForExitAsync(timeout.Token);
        }
        catch (OperationCanceledException exception) when (timeout.IsCancellationRequested)
        {
            if (!process.HasExited)
            {
                try { _terminate(process); }
                catch (InvalidOperationException) when (process.HasExited)
                {
                    // The child exited between the state check and termination request.
                }
            }
            await process.WaitForExitAsync();
            throw new TimeoutException($"{executable} exceeded its hard deadline.\n{await output}\n{await error}", exception);
        }
        return new ProcessResult(process.ExitCode, await output, await error);
    }

    public async Task<ProcessResult> CheckedAsync(
        string executable, IEnumerable<string> arguments, TimeSpan deadline,
        IReadOnlyDictionary<string, string>? environment = null, bool preserveGitContext = false)
    {
        var result = await RunAsync(executable, arguments, deadline, environment, preserveGitContext);
        if (result.ExitCode != 0)
        {
            throw new InvalidOperationException($"{executable} failed with exit code {result.ExitCode}.\n{result.Output}\n{result.Error}");
        }
        return result;
    }
}
