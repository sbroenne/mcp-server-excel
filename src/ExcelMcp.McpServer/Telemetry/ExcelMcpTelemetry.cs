// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Globalization;
using System.Diagnostics;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Channel;
using Microsoft.ApplicationInsights.DataContracts;

namespace Sbroenne.ExcelMcp.McpServer.Telemetry;

/// <summary>
/// Centralized telemetry helper for ExcelMcp MCP Server.
/// Provides usage tracking and performance metrics via Application Insights SDK.
///
/// Telemetry types:
/// - TrackEvent: Tool usage analytics (which tools, actions, success/failure rates)
/// - TrackRequest: Performance metrics (duration, response codes for Performance blade)
/// - TrackException: Unhandled exceptions (for Failures blade)
///
/// User/Session context is applied directly to each telemetry item before sending.
/// View data in Azure Portal: Logs blade with Kusto queries on customEvents/requests tables.
/// </summary>
public static class ExcelMcpTelemetry
{
    private const string RedactedExceptionMessage = "[REDACTED]";
    private const int MaxExceptionDepth = 16;

    /// <summary>
    /// Unique session ID for correlating telemetry within a single MCP server process.
    /// Changes each time the MCP server starts.
    /// </summary>
    public static readonly string SessionId = Guid.NewGuid().ToString("N")[..8];

    /// <summary>
    /// Stable anonymous user ID based on machine identity.
    /// Persists across sessions for the same machine, enabling user-level analytics
    /// without collecting personally identifiable information.
    /// </summary>
    public static readonly string UserId = GenerateAnonymousUserId();

    /// <summary>
    /// Application Insights TelemetryClient for sending Custom Events.
    /// Enables Users/Sessions analytics in Azure Portal.
    /// </summary>
    private static TelemetryClient? _telemetryClient;

    /// <summary>
    /// Sets the TelemetryClient instance for sending Custom Events.
    /// Called by Program.cs during startup. Also tracks a session start event.
    /// </summary>
    internal static void SetTelemetryClient(TelemetryClient client)
    {
        _telemetryClient = client;

        // Track session start to ensure Users/Sessions blades have data
        // This event fires once per MCP server process startup
        TrackSessionStart();
    }

    /// <summary>
    /// Tracks the start of an MCP server session.
    /// Uses an EventTelemetry instance so user/session/version context can be applied explicitly.
    /// This ensures Users/Sessions blades have data even if no tools are invoked.
    /// </summary>
    private static void TrackSessionStart()
    {
        if (_telemetryClient == null) return;

        var telemetry = new EventTelemetry("SessionStart");
        telemetry.Properties["SessionId"] = SessionId;
        telemetry.Properties["AppVersion"] = GetVersion();
        ApplyContext(telemetry);
        _telemetryClient.TrackEvent(telemetry);
    }

    /// <summary>
    /// Flushes any buffered telemetry to Application Insights.
    /// CRITICAL: Must be called before application exits to ensure telemetry is not lost.
    /// Application Insights SDK buffers telemetry and sends in batches - without explicit flush,
    /// short-lived processes like MCP servers may terminate before telemetry is transmitted.
    /// </summary>
    public static void Flush()
    {
        if (_telemetryClient == null) return;

        try
        {
            // Flush with timeout to avoid hanging on shutdown
            // 2 seconds is sufficient; longer waits risk stalling on DNS/network failures
            _telemetryClient.FlushAsync(CancellationToken.None).Wait(TimeSpan.FromSeconds(2));
        }
        catch (Exception)
        {
            // Don't let telemetry flush failure crash the application
        }
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
    /// - TrackEvent: For tool usage analytics (customEvents table)
    /// - TrackRequest: For performance metrics (requests table, Performance blade)
    /// </summary>
    /// <param name="toolName">The MCP tool name (e.g., "range")</param>
    /// <param name="action">The action performed (e.g., "get-values")</param>
    /// <param name="durationMs">Duration in milliseconds</param>
    /// <param name="success">Whether the operation succeeded</param>
    /// <param name="excelPath">Optional Excel file path (will be hashed for privacy)</param>
    public static void TrackToolInvocation(string toolName, string action, long durationMs, bool success, string? excelPath = null)
    {
        if (_telemetryClient == null) return;

        var operationName = $"{toolName}/{action}";
        var startTime = DateTimeOffset.UtcNow.AddMilliseconds(-durationMs);
        var duration = TimeSpan.FromMilliseconds(durationMs);

        var properties = new Dictionary<string, string>
        {
            ["Tool"] = toolName,
            ["Action"] = action,
            ["Success"] = success.ToString()
        };

        // Add hashed file path for grouping (if provided)
        if (!string.IsNullOrEmpty(excelPath))
        {
            properties["FileSessionId"] = HashFilePath(excelPath);
        }

        // Track as customEvent for analytics (tool usage, parameters, success/failure)
        var eventTelemetry = new EventTelemetry(operationName);
        foreach (var property in properties)
        {
            eventTelemetry.Properties[property.Key] = property.Value;
        }
        eventTelemetry.Properties["DurationMs"] = durationMs.ToString(CultureInfo.InvariantCulture);

        ApplyContext(eventTelemetry);
        _telemetryClient.TrackEvent(eventTelemetry);

        // Track as request for Performance blade, Failures blade, Smart Detection
        var request = new RequestTelemetry
        {
            Name = operationName,
            Timestamp = startTime,
            Duration = duration,
            ResponseCode = success ? "200" : "500",
            Success = success
        };

        // Copy properties to request for consistent filtering
        foreach (var prop in properties)
        {
            request.Properties[prop.Key] = prop.Value;
        }

        ApplyContext(request);
        _telemetryClient.TrackRequest(request);
    }

    /// <summary>
    /// Tracks an unhandled exception.
    /// Only call this for exceptions that escape all catch blocks (true bugs/crashes).
    /// </summary>
    /// <param name="exception">The unhandled exception</param>
    /// <param name="source">Source of the exception (e.g., "AppDomain.UnhandledException")</param>
    public static void TrackUnhandledException(Exception exception, string source)
    {
        if (_telemetryClient == null || exception == null) return;

        var telemetry = CreateSanitizedExceptionTelemetry(exception, source);
        _telemetryClient.TrackException(telemetry);
    }

    internal static ExceptionTelemetry CreateSanitizedExceptionTelemetry(
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
                    .Order(StringComparer.Ordinal)),
            ["AppVersion"] = GetVersion()
        };

        if (failureSite != null)
        {
            properties["FailureSite"] = failureSite;
        }

        var details = exceptions.Select((item, index) =>
            new ExceptionDetailsInfo(
                id: index + 1,
                outerId: index == 0 ? 0 : 1,
                typeName: item.GetType().FullName ?? item.GetType().Name,
                message: RedactedExceptionMessage,
                hasFullStack: false,
                stack: string.Empty,
                parsedStack: Array.Empty<Microsoft.ApplicationInsights.DataContracts.StackFrame>()));
        var problemId = failureSite == null
            ? $"{exceptionType} at {safeSource}"
            : $"{exceptionType} at {failureSite}";
        var telemetry = new ExceptionTelemetry(
            details,
            SeverityLevel.Critical,
            problemId,
            properties);
        ApplyContext(telemetry);
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

    /// <summary>
    /// Gets the application version from assembly metadata.
    /// </summary>
    private static string GetVersion()
    {
        return Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
            ?? Assembly.GetExecutingAssembly().GetName().Version?.ToString()
            ?? "1.0.0";
    }

    /// <summary>
    /// Generates a stable anonymous user ID based on machine identity.
    /// Uses a hash of machine name and user profile path to create a consistent
    /// identifier that persists across sessions without collecting PII.
    /// </summary>
    private static string GenerateAnonymousUserId()
    {
        try
        {
            // Combine machine-specific values that are stable but not personally identifiable
            var machineIdentity = $"{Environment.MachineName}|{Environment.UserName}|{Environment.OSVersion.Platform}";

            // Create a SHA256 hash and take the first 16 characters
            var bytes = Encoding.UTF8.GetBytes(machineIdentity);
            var hash = SHA256.HashData(bytes);
            return Convert.ToHexString(hash)[..16].ToLowerInvariant();
        }
        catch (Exception)
        {
            // Fallback to a random ID if machine identity cannot be determined
            return Guid.NewGuid().ToString("N")[..16];
        }
    }

    /// <summary>
    /// Hashes a file path for privacy-preserving grouping.
    /// Enables grouping telemetry by file without exposing actual file paths.
    /// </summary>
    /// <param name="filePath">The file path to hash</param>
    /// <returns>First 12 characters of SHA256 hash (lowercase hex)</returns>
    private static string HashFilePath(string filePath)
    {
        var bytes = Encoding.UTF8.GetBytes(filePath.ToLowerInvariant());
        var hash = SHA256.HashData(bytes);
        return Convert.ToHexString(hash)[..12].ToLowerInvariant();
    }

    private static void ApplyContext(ITelemetry telemetry)
    {
        telemetry.Context.User.Id ??= UserId;
        telemetry.Context.Session.Id ??= SessionId;
        telemetry.Context.Cloud.RoleName ??= "ExcelMcp.McpServer";
        telemetry.Context.Cloud.RoleInstance = $"instance-{UserId[..8]}";
        telemetry.Context.Component.Version = GetVersion();
    }
}

