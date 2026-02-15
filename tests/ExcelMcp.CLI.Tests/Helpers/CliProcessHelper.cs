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
        // The CLI project is a project reference, so the exe is in the same output directory
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
    /// <param name="args">Arguments to pass to excelcli (e.g., "diag ping")</param>
    /// <param name="timeoutMs">Timeout in milliseconds (default: 30000)</param>
    /// <param name="environmentVariables">Optional environment variables to set on the process</param>
    /// <returns>Process result with stdout, stderr, and exit code</returns>
    public static async Task<CliResult> RunAsync(string args, int timeoutMs = 30000, Dictionary<string, string>? environmentVariables = null)
    {
        var exePath = GetExePath();
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
            throw new TimeoutException($"excelcli timed out after {timeoutMs}ms. Args: {args}");
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
        string args, int timeoutMs = 30000, Dictionary<string, string>? environmentVariables = null)
    {
        var result = await RunAsync(args, timeoutMs, environmentVariables);
        var json = JsonDocument.Parse(result.Stdout);
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
