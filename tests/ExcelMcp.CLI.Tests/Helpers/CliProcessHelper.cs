using System.Diagnostics;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.CLI.Tests.Helpers;

/// <summary>
/// Helper for running excelcli as a subprocess and capturing output.
/// Used by integration tests that verify CLI behavior end-to-end.
/// </summary>
internal static class CliProcessHelper
{
    /// <summary>
    /// Gets the path to the excelcli executable.
    /// Finds it relative to the test assembly location.
    /// </summary>
    public static string GetExePath()
    {
        var testDir = AppContext.BaseDirectory;
        var exePath = Path.Combine(testDir, "excelcli.exe");

        if (!File.Exists(exePath))
        {
            throw new FileNotFoundException(
                $"excelcli.exe not found at {exePath}. Ensure ExcelMcp.CLI is a project reference.");
        }

        return exePath;
    }

    /// <summary>
    /// Runs an excelcli command and captures the result.
    /// Always uses -q (quiet) mode for clean JSON output.
    /// </summary>
    public static async Task<CliResult> RunAsync(
        string args,
        int timeoutMs = 30000,
        Dictionary<string, string>? environmentVariables = null,
        string? diagnosticLabel = null)
    {
        var exePath = GetExePath();
        var commandLabel = string.IsNullOrWhiteSpace(diagnosticLabel) ? args : diagnosticLabel;
        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = $"-q {args}",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = Path.GetDirectoryName(exePath)!
        };

        if (environmentVariables != null)
        {
            foreach (var (key, value) in environmentVariables)
            {
                startInfo.Environment[key] = value;
            }
        }

        using var process = new Process { StartInfo = startInfo };
        var stdout = new System.Text.StringBuilder();
        var stderr = new System.Text.StringBuilder();

        process.OutputDataReceived += (_, e) => { if (e.Data != null) stdout.AppendLine(e.Data); };
        process.ErrorDataReceived += (_, e) => { if (e.Data != null) stderr.AppendLine(e.Data); };

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        var completed = await process.WaitForExitAsync(new CancellationTokenSource(timeoutMs).Token)
            .ContinueWith(t => !t.IsCanceled);

        if (!completed)
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"excelcli command '{commandLabel}' timed out after {timeoutMs}ms.");
        }

        return new CliResult
        {
            ExitCode = process.ExitCode,
            Stdout = stdout.ToString().Trim(),
            Stderr = stderr.ToString().Trim()
        };
    }

    /// <summary>
    /// Runs an excelcli command from discrete arguments and captures the result.
    /// Uses ProcessStartInfo.ArgumentList so tests can pass JSON and paths without manual quoting.
    /// </summary>
    public static async Task<CliResult> RunAsync(
        IReadOnlyList<string> args,
        int timeoutMs = 30000,
        Dictionary<string, string>? environmentVariables = null,
        string? diagnosticLabel = null)
    {
        var exePath = GetExePath();
        var commandLabel = string.IsNullOrWhiteSpace(diagnosticLabel) ? string.Join(" ", args) : diagnosticLabel;
        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = Path.GetDirectoryName(exePath)!
        };

        startInfo.ArgumentList.Add("-q");
        foreach (var arg in args)
        {
            startInfo.ArgumentList.Add(arg);
        }

        if (environmentVariables != null)
        {
            foreach (var (key, value) in environmentVariables)
            {
                startInfo.Environment[key] = value;
            }
        }

        using var process = new Process { StartInfo = startInfo };
        var stdout = new System.Text.StringBuilder();
        var stderr = new System.Text.StringBuilder();

        process.OutputDataReceived += (_, e) => { if (e.Data != null) stdout.AppendLine(e.Data); };
        process.ErrorDataReceived += (_, e) => { if (e.Data != null) stderr.AppendLine(e.Data); };

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        var completed = await process.WaitForExitAsync(new CancellationTokenSource(timeoutMs).Token)
            .ContinueWith(t => !t.IsCanceled);

        if (!completed)
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"excelcli command '{commandLabel}' timed out after {timeoutMs}ms.");
        }

        return new CliResult
        {
            ExitCode = process.ExitCode,
            Stdout = stdout.ToString().Trim(),
            Stderr = stderr.ToString().Trim()
        };
    }

    /// <summary>
    /// Runs an excelcli command and parses the JSON output.
    /// </summary>
    public static async Task<(CliResult Result, JsonDocument Json)> RunJsonAsync(
        string args,
        int timeoutMs = 30000,
        Dictionary<string, string>? environmentVariables = null,
        string? diagnosticLabel = null)
    {
        var commandLabel = string.IsNullOrWhiteSpace(diagnosticLabel) ? args : diagnosticLabel;
        var result = await RunAsync(args, timeoutMs, environmentVariables, diagnosticLabel);
        JsonDocument json;
        try
        {
            json = JsonDocument.Parse(result.Stdout);
        }
        catch (JsonException ex)
        {
            throw new InvalidOperationException(
                $"excelcli command '{commandLabel}' returned invalid JSON. Stdout: {result.Stdout}{Environment.NewLine}Stderr: {result.Stderr}",
                ex);
        }

        return (result, json);
    }

    /// <summary>
    /// Runs an excelcli command from discrete arguments and parses the JSON output.
    /// </summary>
    public static async Task<(CliResult Result, JsonDocument Json)> RunJsonAsync(
        IReadOnlyList<string> args,
        int timeoutMs = 30000,
        Dictionary<string, string>? environmentVariables = null,
        string? diagnosticLabel = null)
    {
        var commandLabel = string.IsNullOrWhiteSpace(diagnosticLabel) ? string.Join(" ", args) : diagnosticLabel;
        var result = await RunAsync(args, timeoutMs, environmentVariables, diagnosticLabel);
        JsonDocument json;
        try
        {
            json = JsonDocument.Parse(result.Stdout);
        }
        catch (JsonException ex)
        {
            throw new InvalidOperationException(
                $"excelcli command '{commandLabel}' returned invalid JSON. Stdout: {result.Stdout}{Environment.NewLine}Stderr: {result.Stderr}",
                ex);
        }

        return (result, json);
    }
}

/// <summary>
/// Result of running excelcli as a subprocess.
/// </summary>
internal sealed class CliResult
{
    public int ExitCode { get; init; }
    public string Stdout { get; init; } = string.Empty;
    public string Stderr { get; init; } = string.Empty;
}
