// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Globalization;
using System.Diagnostics;
using Sbroenne.ExcelMcp.Telemetry;

namespace Sbroenne.ExcelMcp.McpServer.Telemetry;

/// <summary>Privacy-safe MCP usage reporting through the shared batch pipeline.</summary>
public static class ExcelMcpTelemetry
{
    private const int MaxExceptionDepth = 16;
    private static UsageTelemetry _reporter = new("ExcelMcp.McpServer");

    internal static void Start(TelemetrySettings settings) => _reporter.Start(settings);
    internal static void SignalExit() => _reporter.SignalExit();

    internal static IDisposable UseReporterForTesting(UsageTelemetry reporter)
    {
        var previous = _reporter;
        _reporter = reporter;
        return new ReporterScope(() => _reporter = previous);
    }

    private sealed class ReporterScope(Action restore) : IDisposable
    {
        public void Dispose() => restore();
    }

    /// <summary>
    /// Gets the Application Insights connection string (embedded at build time).
    /// </summary>
    public static string? GetConnectionString()
    {
        // Connection string is embedded at build time from Directory.Build.props.user
        // Returns null if not set (placeholder value starts with __)
        if (string.IsNullOrEmpty(TelemetryConfig.ConnectionString) ||
            TelemetryConfig.ConnectionString.StartsWith("__", StringComparison.Ordinal))
        {
            return null;
        }
        return TelemetryConfig.ConnectionString;
    }

    /// <summary>
    /// Tracks a tool invocation with usage and performance metrics.
    /// Uses a custom event and operation trace containing only approved metadata.
    /// </summary>
    /// <param name="toolName">The MCP tool name (e.g., "range")</param>
    /// <param name="action">The action performed (e.g., "get-values")</param>
    /// <param name="durationMs">Duration in milliseconds</param>
    /// <param name="result">Privacy-safe outcome and optional failure classification.</param>
    internal static void TrackToolInvocation(
        string toolName,
        string action,
        long durationMs,
        ToolInvocationResult result)
    {
        if (!_reporter.IsReady) return;

        var (eventTelemetry, requestTelemetry) =
            CreateToolInvocationTelemetry(toolName, action, durationMs, result);
        _reporter.RecordEvent(eventTelemetry);
        _reporter.RecordOperation(requestTelemetry);
    }

    internal static (UsageEvent Event, UsageOperation Request)
        CreateToolInvocationTelemetry(
            string toolName,
            string action,
            long durationMs,
            ToolInvocationResult result)
    {
        var operationName = $"{toolName}/{action}";
        var startTime = DateTimeOffset.UtcNow.AddMilliseconds(-durationMs);
        var duration = TimeSpan.FromMilliseconds(durationMs);
        var requestSucceeded = result.Outcome != ToolInvocationOutcome.Failed;
        var properties = new Dictionary<string, string>
        {
            ["Tool"] = toolName,
            ["Action"] = action,
            ["EntryPoint"] = "mcp-server",
            ["Success"] = requestSucceeded.ToString(),
            ["Outcome"] = GetOutcomeValue(result.Outcome)
        };

        if (result.FailureClass.HasValue)
        {
            properties["FailureClass"] = GetFailureClassValue(result.FailureClass.Value);
        }

        if (result.FailureCause.HasValue)
        {
            properties["FailureCause"] = GetFailureCauseValue(result.FailureCause.Value);
        }

        var eventTelemetry = new UsageEvent(operationName);
        foreach (var property in properties)
        {
            eventTelemetry.Properties[property.Key] = property.Value;
        }
        eventTelemetry.Properties["DurationMs"] = durationMs.ToString(CultureInfo.InvariantCulture);

        var request = new UsageOperation
        {
            Name = operationName,
            Timestamp = startTime,
            Duration = duration,
            Success = requestSucceeded
        };

        foreach (var prop in properties)
        {
            request.Properties[prop.Key] = prop.Value;
        }

        return (eventTelemetry, request);
    }

    private static string GetOutcomeValue(ToolInvocationOutcome outcome) =>
        outcome switch
        {
            ToolInvocationOutcome.Succeeded => "succeeded",
            ToolInvocationOutcome.ExpectedNegative => "expected-negative",
            ToolInvocationOutcome.Failed => "failed",
            _ => throw new ArgumentOutOfRangeException(nameof(outcome), outcome, null)
        };

    private static string GetFailureClassValue(ToolFailureClass failureClass) =>
        failureClass switch
        {
            ToolFailureClass.InputState => "input-state",
            ToolFailureClass.ExternalDependency => "external-dependency",
            ToolFailureClass.TimeoutCancellation => "timeout-cancellation",
            ToolFailureClass.ExcelRuntime => "excel-runtime",
            ToolFailureClass.InternalProductFault => "internal-product-fault",
            ToolFailureClass.Unclassified => "unclassified",
            _ => throw new ArgumentOutOfRangeException(
                nameof(failureClass),
                failureClass,
                null)
        };

    private static string GetFailureCauseValue(ToolFailureCause failureCause) =>
        failureCause switch
        {
            ToolFailureCause.Timeout => "timeout",
            ToolFailureCause.Cancellation => "cancellation",
            _ => throw new ArgumentOutOfRangeException(
                nameof(failureCause),
                failureCause,
                null)
        };

    /// <summary>
    /// Tracks an unhandled exception.
    /// Only call this for exceptions that escape all catch blocks (true bugs/crashes).
    /// </summary>
    /// <param name="exception">The unhandled exception</param>
    /// <param name="source">Source of the exception (e.g., "AppDomain.UnhandledException")</param>
    public static void TrackUnhandledException(Exception exception, string source)
    {
        if (!_reporter.IsReady || exception == null) return;

        var telemetry = CreateSanitizedExceptionTelemetry(exception, source);
        _reporter.RecordEvent(telemetry);
    }

    internal static UsageEvent CreateSanitizedExceptionTelemetry(
        Exception exception,
        string source)
    {
        ArgumentNullException.ThrowIfNull(exception);

        var exceptions = EnumerateExceptions(exception)
            .Take(MaxExceptionDepth)
            .ToArray();
        var exceptionType = exception.GetType().Name;
        var safeSource = NormalizeExceptionSource(source);
        var failureSite = FindOwnedFailureSite(exceptions);
        var properties = new Dictionary<string, string>
        {
            ["Sanitized"] = bool.TrueString.ToLowerInvariant(),
            ["Source"] = safeSource,
            ["ExceptionType"] = exceptionType,
            ["InnerExceptionTypes"] = string.Join(
                ",",
                exceptions
                    .Skip(1)
                    .Select(item => item.GetType().Name)
                    .Distinct(StringComparer.Ordinal)
                    .Order(StringComparer.Ordinal))
        };

        if (failureSite != null)
        {
            properties["FailureSite"] = failureSite;
        }

        var telemetry = new UsageEvent("UnhandledException");
        foreach (var property in properties) telemetry.Properties[property.Key] = property.Value;
        return telemetry;
    }

    private static IEnumerable<Exception> EnumerateExceptions(Exception exception)
    {
        yield return exception;

        if (exception is AggregateException aggregate)
        {
            foreach (var innerException in aggregate.Flatten().InnerExceptions)
            {
                foreach (var nestedException in EnumerateExceptions(innerException))
                {
                    yield return nestedException;
                }
            }

            yield break;
        }

        if (exception.InnerException != null)
        {
            foreach (var innerException in EnumerateExceptions(exception.InnerException))
            {
                yield return innerException;
            }
        }
    }

    private static string NormalizeExceptionSource(string source) =>
        source switch
        {
            "AppDomain.UnhandledException" => source,
            "TaskScheduler.UnobservedTaskException" => source,
            "McpServer.RunAsync" => source,
            _ => "Unknown"
        };

    private static string? FindOwnedFailureSite(IEnumerable<Exception> exceptions)
    {
        foreach (var exception in exceptions)
        {
            var frames = new StackTrace(exception, false).GetFrames();
            if (frames == null)
            {
                continue;
            }

            foreach (var frame in frames)
            {
                var method = frame.GetMethod();
                var declaringType = method?.DeclaringType?.FullName;
                if (declaringType?.StartsWith("Sbroenne.ExcelMcp.", StringComparison.Ordinal) == true)
                {
                    return $"{declaringType}.{method!.Name}";
                }
            }
        }

        return null;
    }

}
