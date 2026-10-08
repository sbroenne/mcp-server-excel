using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacExcelBackend
{
    private static readonly SemaphoreSlim OpenGate = new(1, 1);
    private readonly Func<ProcessStartInfo, string?, CancellationToken, Task<MacProcessResult>> _runProcess;

    public MacExcelBackend(
        Func<ProcessStartInfo, string?, CancellationToken, Task<MacProcessResult>>? runProcess = null)
    {
        _runProcess = runProcess ?? RunProcessAsync;
    }

    internal async Task<MacHelperInfo> RequireHelperAsync(IReadOnlyCollection<string> requiredPrimitives, TimeSpan timeout)
    {
        var result = await InvokeAsync("helper.check", new { }, timeout);
        if (!result.TryGetProperty("helper", out var helper))
        {
            throw new MacExcelOperationException("HelperMalformed", "Excel did not return a helper handshake.");
        }
        return MacHelperProtocol.Validate(helper.GetRawText(), requiredPrimitives);
    }

    public async Task<JsonElement> InvokeAsync(
        string command,
        object? arguments,
        TimeSpan timeout,
        bool allowFailureResult = false)
    {
        var started = Stopwatch.GetTimestamp();
        using var timeoutCts = new CancellationTokenSource(timeout);
        var serializedArguments = JsonSerializer.Serialize(arguments, ServiceProtocol.JsonOptions);

        TimeSpan RemainingTime()
        {
            timeoutCts.Token.ThrowIfCancellationRequested();
            var remaining = timeout - Stopwatch.GetElapsedTime(started);
            if (remaining <= TimeSpan.Zero)
            {
                timeoutCts.Cancel();
                timeoutCts.Token.ThrowIfCancellationRequested();
            }
            return remaining;
        }

        try
        {
            timeoutCts.Token.ThrowIfCancellationRequested();
            if (command == "session.open")
            {
                await OpenGate.WaitAsync(timeoutCts.Token);
                try
                {
                    await using var crossProcessLock = await AcquireOpenLockAsync(timeoutCts.Token);
                    using var args = JsonDocument.Parse(serializedArguments);
                    var filePath = args.RootElement.GetProperty("filePath").GetString();
                    ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
                    try
                    {
                        await InvokeScriptAsync("session.prepare-open", serializedArguments, RemainingTime(), timeoutCts.Token);
                    }
                    catch (MacExcelOperationException error)
                        when (error.Message.Contains("OSStatus -600", StringComparison.Ordinal))
                    {
                        var launchExcel = new ProcessStartInfo("/usr/bin/open");
                        launchExcel.ArgumentList.Add("-g");
                        launchExcel.ArgumentList.Add("-a");
                        launchExcel.ArgumentList.Add("Microsoft Excel");
                        EnsureSuccessfulExit(
                            "Excel startup",
                            await _runProcess(launchExcel, null, timeoutCts.Token));
                        while (true)
                        {
                            try
                            {
                                await InvokeScriptAsync("session.prepare-open", serializedArguments, RemainingTime(), timeoutCts.Token);
                                break;
                            }
                            catch (MacExcelOperationException startupError)
                                when (startupError.Message.Contains("OSStatus -600", StringComparison.Ordinal))
                            {
                                await Task.Delay(TimeSpan.FromMilliseconds(200), timeoutCts.Token);
                            }
                        }
                    }

                    var handoff = new ProcessStartInfo("/usr/bin/open");
                    handoff.ArgumentList.Add("-g");
                    handoff.ArgumentList.Add("-b");
                    handoff.ArgumentList.Add("com.microsoft.Excel");
                    handoff.ArgumentList.Add(filePath);
                    timeoutCts.Token.ThrowIfCancellationRequested();
                    try
                    {
                        var result = await _runProcess(handoff, null, timeoutCts.Token);
                        EnsureSuccessfulExit("LaunchServices handoff", result);
                        var attached = await InvokeScriptAsync(command, serializedArguments, RemainingTime(), timeoutCts.Token);
                        if (!attached.TryGetProperty("success", out var success)
                            || success.ValueKind != JsonValueKind.True
                            || (attached.TryGetProperty("errorMessage", out var error)
                                && !string.IsNullOrEmpty(error.GetString())))
                        {
                            throw new InvalidDataException("Excel did not confirm successful workbook attachment.");
                        }
                        return attached;
                    }
                    catch (Exception error) when (error is OperationCanceledException
                        or TimeoutException or InvalidOperationException or IOException or InvalidDataException or JsonException or Win32Exception)
                    {
                        throw new MacExcelOperationException(
                            "RecoveryRequired",
                            "The Excel file-open outcome is uncertain; the request may still complete. " +
                            "Do not retry automatically or delete the workbook. Unlock the macOS desktop " +
                            "and resolve any pending Excel dialogs, then reconcile the exact workbook. " +
                            $"Original failure: {error.Message}",
                            error);
                    }
                }
                finally
                {
                    OpenGate.Release();
                }
            }

            return await InvokeScriptAsync(
                command,
                serializedArguments,
                RemainingTime(),
                timeoutCts.Token,
                allowFailureResult);
        }
        catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested)
        {
            throw new TimeoutException(
                $"Mac Excel operation '{command}' exceeded {timeout.TotalSeconds:0.###} seconds. " +
                "The workbook session is no longer safe to use.");
        }
    }

    private async Task<JsonElement> InvokeScriptAsync(
        string command,
        string arguments,
        TimeSpan timeout,
        CancellationToken cancellationToken,
        bool allowFailureResult = false)
    {
        var startInfo = CreateAutomationStartInfo(command, timeout);
        var result = await _runProcess(startInfo, arguments, cancellationToken);
        EnsureSuccessfulExit(command, result);

        using var document = JsonDocument.Parse(result.StandardOutput);
        var root = document.RootElement.Clone();
        if (root.TryGetProperty("success", out var success)
            && !success.GetBoolean()
            && (!allowFailureResult || !root.TryGetProperty("filePath", out _)))
        {
            var message = root.TryGetProperty("errorMessage", out var error)
                ? error.GetString()
                : "Unknown Mac Excel error.";
            var category = root.TryGetProperty("errorCategory", out var errorCategory)
                ? errorCategory.GetString()
                : null;
            if (category == "Timeout")
            {
                throw new TimeoutException(message ?? "Excel Apple Event timed out; its outcome may be uncertain.");
            }
            throw new MacExcelOperationException(
                category ?? "ComInterop",
                message ?? "Unknown Mac Excel error.",
                remoteExceptionType: root.TryGetProperty("exceptionType", out var exceptionType) ? exceptionType.GetString() : null,
                remoteInnerError: root.TryGetProperty("innerError", out var innerError) ? innerError.GetString() : null);
        }

        return root;
    }

    private static ProcessStartInfo CreateAutomationStartInfo(string command, TimeSpan timeout)
    {
        var processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Current process path is unavailable.");
        var startInfo = new ProcessStartInfo(processPath)
        {
            RedirectStandardInput = true
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase)
            && Assembly.GetEntryAssembly()?.GetName().Name is { Length: > 0 } entryAssemblyName)
        {
            var entryAssemblyPath = Path.Combine(AppContext.BaseDirectory, $"{entryAssemblyName}.dll");
            if (!File.Exists(entryAssemblyPath))
            {
                throw new InvalidOperationException(
                    $"Entry assembly '{entryAssemblyPath}' is unavailable for Mac automation.");
            }
            startInfo.ArgumentList.Add(entryAssemblyPath);
        }
        startInfo.ArgumentList.Add(MacAutomationHost.Marker);
        startInfo.ArgumentList.Add(command);
        startInfo.ArgumentList.Add(Environment.ProcessId.ToString(
            System.Globalization.CultureInfo.InvariantCulture));
        startInfo.ArgumentList.Add(timeout.Ticks.ToString(
            System.Globalization.CultureInfo.InvariantCulture));
        return startInfo;
    }

    private static async Task<FileStream> AcquireOpenLockAsync(CancellationToken cancellationToken)
    {
        var lockPath = Path.Combine(
            Path.GetTempPath(),
            $"excelmcp-launchservices-open-{ServiceSecurity.GetCurrentUserIdentity()}.lock");
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                return new FileStream(
                    lockPath,
                    FileMode.OpenOrCreate,
                    FileAccess.ReadWrite,
                    FileShare.None,
                    bufferSize: 1,
                    FileOptions.Asynchronous);
            }
            catch (IOException)
            {
                await Task.Delay(50, cancellationToken);
            }
        }
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

internal sealed class MacExcelOperationException(
    string errorCategory, string message, Exception? innerException = null,
    string? remoteExceptionType = null, string? remoteInnerError = null)
    : InvalidOperationException(message, innerException)
{
    public string ErrorCategory { get; } = errorCategory;
    public string? RemoteExceptionType { get; } = remoteExceptionType;
    public string? RemoteInnerError { get; } = remoteInnerError;

    internal ServiceResponse ToServiceResponse() => new()
    {
        Success = false,
        ErrorCategory = ErrorCategory,
        ErrorMessage = RemoteExceptionType is null ? Message : $"{RemoteExceptionType}: {Message}",
        ExceptionType = RemoteExceptionType ?? GetType().Name,
        InnerError = RemoteInnerError
    };
}
