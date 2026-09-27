using System.Diagnostics;
using System.Reflection;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacExcelBackend
{
    private const string ResourceName = "Sbroenne.ExcelMcp.Service.Mac.MacExcelBridge.js";
    private readonly string _script;
    private readonly Func<ProcessStartInfo, string?, CancellationToken, Task<MacProcessResult>> _runProcess;
    private readonly Func<int> _checkPermission;

    public MacExcelBackend(
        Func<ProcessStartInfo, string?, CancellationToken, Task<MacProcessResult>>? runProcess = null,
        Func<int>? checkPermission = null)
    {
        _runProcess = runProcess ?? RunProcessAsync;
        _checkPermission = checkPermission ?? MacAutomationAccess.Check;
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName)
            ?? throw new InvalidOperationException($"Embedded macOS bridge '{ResourceName}' was not found.");
        using var reader = new StreamReader(stream);
        _script = reader.ReadToEnd();
    }

    public async Task<JsonElement> InvokeAsync(string command, object? arguments, TimeSpan timeout)
    {
        using var timeoutCts = new CancellationTokenSource(timeout);
        var serializedArguments = JsonSerializer.Serialize(arguments, ServiceProtocol.JsonOptions);
        var permission = _checkPermission();
        if (permission != 0)
        {
            throw new MacExcelOperationException("ComInterop",
                $"Mac Excel automation is not ready: {MacAutomationAccess.DescribeStatus(permission)} " +
                $"(OSStatus {permission}). Open licensed Excel and grant Automation permission interactively. " +
                "No permission prompt was requested.");
        }

        try
        {
            timeoutCts.Token.ThrowIfCancellationRequested();
            if (command == "session.open")
            {
                using var args = JsonDocument.Parse(serializedArguments);
                var filePath = args.RootElement.GetProperty("filePath").GetString();
                ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
                await InvokeScriptAsync("session.prepare-open", serializedArguments, timeoutCts.Token);

                var handoff = new ProcessStartInfo("/usr/bin/open");
                handoff.ArgumentList.Add("-g");
                handoff.ArgumentList.Add("-b");
                handoff.ArgumentList.Add("com.microsoft.Excel");
                handoff.ArgumentList.Add(filePath);
                var result = await _runProcess(handoff, null, timeoutCts.Token);
                EnsureSuccessfulExit("LaunchServices handoff", result);
            }

            return await InvokeScriptAsync(command, serializedArguments, timeoutCts.Token);
        }
        catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested)
        {
            throw new TimeoutException(
                $"Mac Excel operation '{command}' exceeded {timeout.TotalSeconds:0.###} seconds. " +
                "The workbook session is no longer safe to use.");
        }
    }

    private async Task<JsonElement> InvokeScriptAsync(
        string command, string arguments, CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo("/usr/bin/osascript")
        { RedirectStandardInput = true };
        startInfo.ArgumentList.Add("-l");
        startInfo.ArgumentList.Add("JavaScript");
        startInfo.ArgumentList.Add("-");
        startInfo.ArgumentList.Add(command);
        startInfo.ArgumentList.Add(arguments);
        var result = await _runProcess(startInfo, _script, cancellationToken);
        EnsureSuccessfulExit(command, result);

        using var document = JsonDocument.Parse(result.StandardOutput);
        var root = document.RootElement.Clone();
        if (root.TryGetProperty("success", out var success) && !success.GetBoolean())
        {
            var message = root.TryGetProperty("errorMessage", out var error)
                ? error.GetString()
                : "Unknown Mac Excel error.";
            var category = root.TryGetProperty("errorCategory", out var errorCategory)
                ? errorCategory.GetString()
                : null;
            throw new MacExcelOperationException(
                category ?? "ComInterop",
                message ?? "Unknown Mac Excel error.");
        }

        return root;
    }

    private static void EnsureSuccessfulExit(string command, MacProcessResult result)
    {
        if (result.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"Mac Excel operation '{command}' failed: {SanitizeError(result.StandardError)}");
        }
    }

    private static async Task<MacProcessResult> RunProcessAsync(
        ProcessStartInfo startInfo, string? input, CancellationToken cancellationToken)
    {
        startInfo.UseShellExecute = false;
        startInfo.RedirectStandardOutput = true;
        startInfo.RedirectStandardError = true;
        using var process = new Process { StartInfo = startInfo };
        cancellationToken.ThrowIfCancellationRequested();
        if (!process.Start())
        {
            throw new InvalidOperationException($"Could not start '{startInfo.FileName}'.");
        }
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        try
        {
            if (input is not null)
            {
                await process.StandardInput.WriteAsync(input.AsMemory(), cancellationToken);
                process.StandardInput.Close();
            }
            await process.WaitForExitAsync(cancellationToken);
        }
        catch (OperationCanceledException)
        {
            if (!process.HasExited)
            {
                process.Kill();
            }
            await process.WaitForExitAsync();
            await Task.WhenAll(stdout, stderr);
            throw;
        }
        return new MacProcessResult(process.ExitCode, await stdout, await stderr);
    }

    private static string SanitizeError(string error)
    {
        var text = error.Trim();
        var marker = text.LastIndexOf("execution error:", StringComparison.OrdinalIgnoreCase);
        return marker >= 0 ? text[(marker + "execution error:".Length)..].Trim() : text;
    }
}

internal sealed record MacProcessResult(int ExitCode, string StandardOutput, string StandardError);

internal sealed class MacExcelOperationException(string errorCategory, string message)
    : InvalidOperationException(message)
{
    public string ErrorCategory { get; } = errorCategory;
}
