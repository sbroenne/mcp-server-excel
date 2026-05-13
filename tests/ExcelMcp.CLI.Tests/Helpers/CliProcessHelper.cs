using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;

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

        return await RunProcessAsync(startInfo, commandLabel, timeoutMs, environmentVariables);
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

        return await RunProcessAsync(startInfo, commandLabel, timeoutMs, environmentVariables);
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
                $"excelcli command '{commandLabel}' returned invalid JSON.{Environment.NewLine}{BuildDiagnosticMessage(commandLabel, result, environmentVariables)}",
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
                $"excelcli command '{commandLabel}' returned invalid JSON.{Environment.NewLine}{BuildDiagnosticMessage(commandLabel, result, environmentVariables)}",
                ex);
        }

        return (result, json);
    }

    public static string DescribeDaemonState(Dictionary<string, string>? environmentVariables = null)
    {
        var pipeName = GetPipeName(environmentVariables);
        var trackerPath = DaemonProcessTracker.GetTrackingFilePath(pipeName);
        var builder = new StringBuilder();
        builder.AppendLine(CultureInfo.InvariantCulture, $"Pipe: {pipeName}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"Daemon mutex held: {DaemonAutoStart.IsDaemonMutexHeld(pipeName)}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"Tracker path: {trackerPath}");

        if (File.Exists(trackerPath))
        {
            builder.AppendLine(CultureInfo.InvariantCulture, $"Tracker content: {File.ReadAllText(trackerPath)}");
        }
        else
        {
            builder.AppendLine("Tracker content: <missing>");
        }

        if (DaemonProcessTracker.TryGetTrackedProcess(pipeName, out var trackedProcess))
        {
            using (trackedProcess)
            {
                builder.AppendLine(CultureInfo.InvariantCulture, $"Tracked process: PID {trackedProcess.Id}, HasExited={trackedProcess.HasExited}");
            }
        }
        else
        {
            builder.AppendLine("Tracked process: <none>");
        }

        return builder.ToString().TrimEnd();
    }

    private static async Task<CliResult> RunProcessAsync(
        ProcessStartInfo startInfo,
        string commandLabel,
        int timeoutMs,
        Dictionary<string, string>? environmentVariables)
    {
        using var process = new Process { StartInfo = startInfo };
        process.Start();

        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        using var timeoutCts = new CancellationTokenSource(timeoutMs);
        try
        {
            await process.WaitForExitAsync(timeoutCts.Token);
        }
        catch (OperationCanceledException)
        {
            try
            {
                process.Kill(entireProcessTree: true);
            }
            catch (InvalidOperationException)
            {
                // Process exited between timeout and kill.
            }

            await process.WaitForExitAsync(CancellationToken.None);
            var timedOutResult = new CliResult
            {
                ExitCode = process.HasExited ? process.ExitCode : -1,
                Stdout = (await stdoutTask).Trim(),
                Stderr = (await stderrTask).Trim()
            };

            throw new TimeoutException(
                $"excelcli command '{commandLabel}' timed out after {timeoutMs}ms.{Environment.NewLine}" +
                BuildDiagnosticMessage(commandLabel, timedOutResult, environmentVariables));
        }

        return new CliResult
        {
            ExitCode = process.ExitCode,
            Stdout = (await stdoutTask).Trim(),
            Stderr = (await stderrTask).Trim()
        };
    }

    private static string BuildDiagnosticMessage(
        string commandLabel,
        CliResult result,
        Dictionary<string, string>? environmentVariables)
    {
        var builder = new StringBuilder();
        builder.AppendLine(CultureInfo.InvariantCulture, $"Command: {commandLabel}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"Exit code: {result.ExitCode}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"Stdout: {result.Stdout}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"Stderr: {result.Stderr}");
        builder.AppendLine(DescribeDaemonState(environmentVariables));
        return builder.ToString().TrimEnd();
    }

    private static string GetPipeName(Dictionary<string, string>? environmentVariables)
    {
        if (environmentVariables != null
            && environmentVariables.TryGetValue("EXCELMCP_CLI_PIPE", out var pipeName)
            && !string.IsNullOrWhiteSpace(pipeName))
        {
            return pipeName;
        }

        return DaemonAutoStart.GetPipeName();
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
