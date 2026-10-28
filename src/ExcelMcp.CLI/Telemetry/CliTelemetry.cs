using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Telemetry;

namespace Sbroenne.ExcelMcp.CLI.Telemetry;

internal static class CliTelemetry
{
    private const string EntryPoint = "cli";
    internal const string UnknownOperationPart = "other";
    private const string UnknownCommand = $"{UnknownOperationPart}.{UnknownOperationPart}";

    /// <summary>Commands handled by the service host instead of a generated category.</summary>
    private static readonly Dictionary<string, string[]> BuiltInActions = new(StringComparer.OrdinalIgnoreCase)
    {
        ["service"] = ["start", "stop", "status", "shutdown", "ping"],
        ["session"] = ["create", "open", "save", "close", "list", "test", "discover", "attach",
            "notices", "enable-editing", "enable-connections", "enable-macros"],
        ["batch"] = ["run"]
    };

    private static readonly string[] HelpFlags = ["--help", "-h"];
    private static readonly AsyncLocal<InvocationTelemetryState?> CurrentInvocationTelemetry = new();
    private static UsageTelemetry _reporter = new("ExcelMcp.CLI");
    private static TelemetrySettings? _settings;

    internal static void Enable() => _settings = TelemetrySettings.Read(TelemetryConfig.ConnectionString);

    internal static void StartDaemonReporting()
    {
        if (_settings != null) _reporter.Start(_settings);
    }

    internal static Task<ServiceResponse> TrackDaemonRequestAsync(
        ServiceRequest request, Func<Task<ServiceResponse>> operation) =>
        string.Equals(request.Command, "service.ping", StringComparison.OrdinalIgnoreCase)
            ? operation()
            : TrackCommandAsync(request, operation);

    internal static IDisposable UseReporterForTesting(UsageTelemetry reporter)
    {
        var previous = _reporter;
        var previousSettings = _settings;
        _reporter = reporter;
        return new ReporterScope(() => { _reporter = previous; _settings = previousSettings; });
    }

    private sealed class ReporterScope(Action restore) : IDisposable
    {
        public void Dispose() => restore();
    }

    internal static async Task<ServiceResponse> TrackCommandAsync(
        ServiceRequest request,
        Func<Task<ServiceResponse>> operation) =>
        await TrackCommandAsync(request, operation, TrackCommandInvocation);

    internal static async Task<ServiceResponse> TrackCommandAsync(
        ServiceRequest request,
        Func<Task<ServiceResponse>> operation,
        Action<string, long, bool, string?, bool> trackInvocation)
    {
        var stopwatch = Stopwatch.StartNew();
        ServiceResponse? response = null;
        string? failureCategory = null;
        try
        {
            response = await operation();
            return response;
        }
        catch (Exception ex)
        {
            // Transport failures never reach a response, so classify the thrown
            // exception with the shared classifier instead of losing the reason.
            failureCategory = OperationFailureClassifier.Classify(ex);
            throw;
        }
        finally
        {
            stopwatch.Stop();
            var invocationTelemetry = CurrentInvocationTelemetry.Value;
            string? resultErrorCategory = null;
            var negativeResult = response?.Success == true &&
                ServiceResultOutcome.TryReadNegative(response.Result, out _, out resultErrorCategory, request.Command);
            var expectedNegative = IsExpectedNegative(request, response);
            var succeeded = response?.Success == true && (!negativeResult || expectedNegative);
            var trackedFailureCategory = response?.ErrorCategory
                ?? (expectedNegative ? null : resultErrorCategory)
                ?? failureCategory;
            if (invocationTelemetry != null && response?.Success != true)
            {
                invocationTelemetry.FailureCategory ??= trackedFailureCategory;
            }

            if (invocationTelemetry?.TrackRequests is not false)
            {
                if (invocationTelemetry != null)
                {
                    invocationTelemetry.RequestTracked = true;
                }
                trackInvocation(
                    request.Command,
                    stopwatch.ElapsedMilliseconds,
                    succeeded,
                    trackedFailureCategory,
                    expectedNegative);
            }
        }
    }

    private static bool IsExpectedNegative(
        ServiceRequest request,
        ServiceResponse? response)
    {
        if (response?.Success != true
            || !string.Equals(request.Command, "session.test", StringComparison.OrdinalIgnoreCase)
            || string.IsNullOrWhiteSpace(response.Result))
        {
            return false;
        }

        try
        {
            return ServiceProtocol.Deserialize<FileValidationInfo>(response.Result) is { CanOpen: false };
        }
        catch (JsonException)
        {
            return false;
        }
    }

    internal static void RecordFinalFailure(string errorCategory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(errorCategory);
        var invocationTelemetry = CurrentInvocationTelemetry.Value;
        if (invocationTelemetry != null)
        {
            invocationTelemetry.FailureCategory = errorCategory;
        }
    }

    internal static void RecordExpectedNegative()
    {
        var invocationTelemetry = CurrentInvocationTelemetry.Value;
        if (invocationTelemetry != null)
        {
            invocationTelemetry.ExpectedNegative = true;
            invocationTelemetry.FailureCategory = null;
        }
    }

    internal static void TrackLocalFailure(
        string command,
        long durationMs,
        string errorCategory) =>
        TrackLocalFailure(command, durationMs, errorCategory, TrackCommandInvocation);

    internal static void TrackLocalFailure(
        string command,
        long durationMs,
        string errorCategory,
        Action<string, long, bool, string?, bool> trackInvocation)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(command);
        ArgumentException.ThrowIfNullOrWhiteSpace(errorCategory);
        ArgumentNullException.ThrowIfNull(trackInvocation);

        var invocationTelemetry = CurrentInvocationTelemetry.Value;
        if (invocationTelemetry != null)
        {
            invocationTelemetry.RequestTracked = true;
            invocationTelemetry.FailureCategory ??= errorCategory;
        }

        trackInvocation(command, durationMs, false, errorCategory, false);
    }

    /// <summary>
    /// Tracks the whole CLI invocation so commands that never issue a service
    /// request, and failures raised before one is sent, are still measured.
    /// Regular commands emit their final outcome; batches retain per-item telemetry.
    /// </summary>
    internal static int TrackCliInvocation(string[] args, Func<int> operation)
    {
        return TrackCliInvocation(args, operation, TrackCommandInvocation);
    }

    internal static int TrackCliInvocation(
        string[] args,
        Func<int> operation,
        Action<string, long, bool, string?, bool> trackInvocation) =>
        TrackCliInvocationAsync(args, () => Task.FromResult(operation()), trackInvocation).GetAwaiter().GetResult();

    internal static Task<int> TrackCliInvocationAsync(string[] args, Func<Task<int>> operation)
    {
        return TrackCliInvocationAsync(args, operation, TrackCommandInvocation);
    }

    internal static async Task<int> TrackCliInvocationAsync(
        string[] args,
        Func<Task<int>> operation,
        Action<string, long, bool, string?, bool> trackInvocation)
    {
        if (IsHelpRequest(args))
        {
            return await operation();
        }

        var isBatch = IsBatchCommand(args);
        var previousInvocationTelemetry = CurrentInvocationTelemetry.Value;
        var invocationTelemetry = new InvocationTelemetryState(isBatch);
        CurrentInvocationTelemetry.Value = invocationTelemetry;
        var stopwatch = Stopwatch.StartNew();
        var exitCode = 1;
        string? failureCategory = null;
        var operationThrew = false;
        try
        {
            exitCode = await operation();
            return exitCode;
        }
        catch (Exception ex)
        {
            operationThrew = true;
            failureCategory = OperationFailureClassifier.Classify(ex);
            throw;
        }
        finally
        {
            stopwatch.Stop();
            CurrentInvocationTelemetry.Value = previousInvocationTelemetry;
            var unhandledBatchFailure = isBatch && operationThrew;
            if (!isBatch || !invocationTelemetry.RequestTracked || unhandledBatchFailure)
            {
                var expectedNegative = invocationTelemetry.ExpectedNegative;
                trackInvocation(
                    ResolveCliCommand(args),
                    stopwatch.ElapsedMilliseconds,
                    exitCode == 0 || expectedNegative,
                    unhandledBatchFailure
                        ? failureCategory
                        : invocationTelemetry.FailureCategory ?? failureCategory,
                    expectedNegative);
            }
        }
    }

    internal static bool IsHelpRequest(string[] args) =>
        args.Any(arg => HelpFlags.Contains(arg, StringComparer.OrdinalIgnoreCase));

    private static bool IsBatchCommand(string[] args) =>
        args.Length > 0 && string.Equals(args[0], "batch", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Maps parsed CLI arguments to a canonical service command. Arguments that
    /// are not an allowlisted command never leave the machine.
    /// </summary>
    internal static string ResolveCliCommand(string[] args)
    {
        if (args.Length == 0)
        {
            return UnknownCommand;
        }

        var category = ServiceRegistry.ValidActionsByCategory.ContainsKey(args[0])
            ? args[0].ToLowerInvariant()
            : args[0];
        // Commands without a subcommand (for example "batch") use the "run" action.
        var action = args.Length > 1 && !args[1].StartsWith('-') ? args[1] : "run";
        return $"{category}.{action}";
    }

    internal static void SignalExit() => _reporter.SignalExit();

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL001",
        Justification = "Optional telemetry construction/reporting must never change command results or exit codes; dropping a record avoids recursive reporting through a failing pipeline.")]
    internal static void TrackCommandInvocation(
        string command,
        long durationMs,
        bool succeeded,
        string? errorCategory,
        bool expectedNegative)
    {
        try
        {
            if (!_reporter.IsReady) return;
            var (eventTelemetry, requestTelemetry) =
                CreateCommandInvocationTelemetry(
                    command,
                    durationMs,
                    succeeded,
                    errorCategory,
                    expectedNegative);
            _reporter.RecordEvent(eventTelemetry);
            _reporter.RecordOperation(requestTelemetry);
        }
        catch (Exception)
        {
        }
    }

    internal static (UsageEvent Event, UsageOperation Request)
        CreateCommandInvocationTelemetry(
            string command,
            long durationMs,
            bool succeeded,
            string? errorCategory,
            bool expectedNegative = false)
    {
        var (tool, action) = ResolveOperation(command);
        var operationName = $"{tool}/{action}";
        var duration = TimeSpan.FromMilliseconds(durationMs);
        var properties = new Dictionary<string, string>
        {
            ["Tool"] = tool,
            ["Action"] = action,
            ["EntryPoint"] = EntryPoint,
            ["Success"] = succeeded.ToString(),
            ["Outcome"] = expectedNegative
                ? "expected-negative"
                : succeeded ? "succeeded" : "failed"
        };

        if (!succeeded)
        {
            properties["FailureClass"] = ClassifyFailure(errorCategory);
            var failureCause = errorCategory switch
            {
                "Timeout" => "timeout",
                "Cancelled" => "cancellation",
                _ => null
            };
            if (failureCause != null)
            {
                properties["FailureCause"] = failureCause;
            }
        }

        var eventTelemetry = new UsageEvent(operationName);
        foreach (var property in properties)
        {
            eventTelemetry.Properties[property.Key] = property.Value;
        }
        eventTelemetry.Properties["DurationMs"] =
            durationMs.ToString(CultureInfo.InvariantCulture);

        var requestTelemetry = new UsageOperation
        {
            Name = operationName,
            Timestamp = DateTimeOffset.UtcNow.Subtract(duration),
            Duration = duration,
            Success = succeeded
        };
        foreach (var property in properties)
        {
            requestTelemetry.Properties[property.Key] = property.Value;
        }

        return (eventTelemetry, requestTelemetry);
    }

    /// <summary>
    /// Maps a service command to allowlisted telemetry parts. Commands outside
    /// the generated and built-in command sets report fixed labels so
    /// user-supplied text is never sent.
    /// </summary>
    internal static (string Tool, string Action) ResolveOperation(string? command)
    {
        var parts = command?.Split('.', 2) ?? [];
        if (parts.Length != 2)
        {
            return (UnknownOperationPart, UnknownOperationPart);
        }

        var category = parts[0];
        var action = parts[1];
        IReadOnlyList<string>? validActions =
            BuiltInActions.TryGetValue(category, out var builtIn) ? builtIn
            : ServiceRegistry.ValidActionsByCategory.TryGetValue(category, out var generated) ? generated
            : null;
        if (validActions == null)
        {
            return (UnknownOperationPart, UnknownOperationPart);
        }

        var canonicalAction = validActions
            .FirstOrDefault(valid => string.Equals(valid, action, StringComparison.OrdinalIgnoreCase));
        return canonicalAction == null
            ? (UnknownOperationPart, UnknownOperationPart)
            : (category.ToLowerInvariant(), canonicalAction);
    }

    private static string ClassifyFailure(string? errorCategory) =>
        errorCategory switch
        {
            "InvalidInput" or "SessionNotFound" or "SessionUnavailable"
                or "NotFound" or "Conflict" or "Syntax" or "Expression" or "Prerequisite"
                => "input-state",
            "Privacy" or "PrivacyFirewall" or "Authentication" or "Connectivity" or "Permissions"
                or "DependencyUnavailable" => "external-dependency",
            "Timeout" or "Cancelled" or "SessionInvalidated" => "timeout-cancellation",
            "ComInterop" or "ExcelProcessDied" or "Cleanup" => "excel-runtime",
            "ServiceUnavailable" or "ServiceStartup" or "InvalidResponse"
                => "internal-product-fault",
            _ => "unclassified"
        };

    private sealed class InvocationTelemetryState(bool trackRequests)
    {
        public bool TrackRequests { get; } = trackRequests;

        public bool RequestTracked { get; set; }

        public string? FailureCategory { get; set; }

        public bool ExpectedNegative { get; set; }
    }
}
