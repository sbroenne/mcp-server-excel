using System.Diagnostics;
using System.Text;

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
        using var timeout = new CancellationTokenSource(deadline);
        var output = new StringBuilder();
        var error = new StringBuilder();
        var capture = Task.WhenAll(
            CaptureAsync(process.StandardOutput, output, timeout.Token),
            CaptureAsync(process.StandardError, error, timeout.Token));
        try
        {
            await Task.WhenAll(process.WaitForExitAsync(timeout.Token), capture).WaitAsync(timeout.Token);
            timeout.Token.ThrowIfCancellationRequested();
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
            using var cleanup = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            try
            {
                await process.WaitForExitAsync(cleanup.Token);
                await capture.WaitAsync(cleanup.Token);
            }
            catch (Exception cleanupException)
            {
                throw new AggregateException($"{executable} exceeded its hard deadline and process/output cleanup failed.",
                    exception, cleanupException);
            }
            throw new TimeoutException($"{executable} exceeded its hard deadline.\n{output}\n{error}", exception);
        }
        return new ProcessResult(process.ExitCode, output.ToString(), error.ToString());
    }

    private static async Task CaptureAsync(StreamReader reader, StringBuilder target, CancellationToken cancellationToken)
    {
        var buffer = new char[4096];
        try
        {
            int count;
            while ((count = await reader.ReadAsync(buffer.AsMemory(), cancellationToken)) != 0)
            {
                target.Append(buffer, 0, count);
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            // The caller reports the deadline with all output captured before cancellation.
        }
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
