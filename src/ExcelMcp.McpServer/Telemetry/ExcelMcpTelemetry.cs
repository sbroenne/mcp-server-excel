// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Diagnostics;
using System.Reflection;

namespace Sbroenne.ExcelMcp.McpServer.Telemetry;

/// <summary>
/// Centralized telemetry helper for ExcelMcp MCP Server.
/// Provides usage tracking and unhandled exception reporting via OpenTelemetry.
/// </summary>
public static class ExcelMcpTelemetry
{
    /// <summary>
    /// The ActivitySource for creating traces.
    /// </summary>
    public static readonly ActivitySource ActivitySource = new("ExcelMcp.McpServer", GetVersion());

    /// <summary>
    /// Environment variable to opt-out of telemetry.
    /// Set to "true" or "1" to disable telemetry.
    /// </summary>
    public const string OptOutEnvironmentVariable = "EXCELMCP_TELEMETRY_OPTOUT";

    /// <summary>
    /// Environment variable to enable debug mode (console output instead of Azure).
    /// Set to "true" or "1" for local testing without Azure resources.
    /// </summary>
    public const string DebugTelemetryEnvironmentVariable = "EXCELMCP_DEBUG_TELEMETRY";

    /// <summary>
    /// Application Insights connection string (embedded at build time).
    /// </summary>
    /// <remarks>
    /// This value is replaced during CI/CD build from the APPINSIGHTS_CONNECTION_STRING secret.
    /// Format: InstrumentationKey=xxx;IngestionEndpoint=https://xxx.in.applicationinsights.azure.com/;...
    /// </remarks>
    private const string ConnectionString = "__APPINSIGHTS_CONNECTION_STRING__";

    /// <summary>
    /// Unique session ID for correlating telemetry within a single MCP server process.
    /// </summary>
    public static readonly string SessionId = Guid.NewGuid().ToString("N")[..8];

    private static bool? _isEnabled;

    /// <summary>
    /// Gets whether telemetry is enabled (not opted out and either debug mode or connection string available).
    /// </summary>
    public static bool IsEnabled
    {
        get
        {
            _isEnabled ??= !IsOptedOut() && (IsDebugMode() || !string.IsNullOrEmpty(GetConnectionString()));
            return _isEnabled.Value;
        }
    }

    /// <summary>
    /// Checks if debug telemetry mode is enabled (console output for local testing).
    /// </summary>
    public static bool IsDebugMode()
    {
        var debug = Environment.GetEnvironmentVariable(DebugTelemetryEnvironmentVariable);
        return string.Equals(debug, "true", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(debug, "1", StringComparison.Ordinal);
    }

    /// <summary>
    /// Gets the Application Insights connection string (embedded at build time).
    /// </summary>
    public static string? GetConnectionString()
    {
        // Connection string is embedded at build time via CI/CD
        // Returns null if placeholder wasn't replaced (local dev builds)
        return ConnectionString.StartsWith("__", StringComparison.Ordinal) ? null : ConnectionString;
    }

    /// <summary>
    /// Checks if user has opted out of telemetry via environment variable.
    /// </summary>
    public static bool IsOptedOut()
    {
        var optOut = Environment.GetEnvironmentVariable(OptOutEnvironmentVariable);
        return string.Equals(optOut, "true", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(optOut, "1", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tracks a tool invocation with usage metrics.
    /// </summary>
    /// <param name="toolName">The MCP tool name (e.g., "excel_range")</param>
    /// <param name="action">The action performed (e.g., "get-values")</param>
    /// <param name="durationMs">Duration in milliseconds</param>
    /// <param name="success">Whether the operation succeeded</param>
    public static void TrackToolInvocation(string toolName, string action, long durationMs, bool success)
    {
        if (!IsEnabled) return;

        using var activity = ActivitySource.StartActivity("ToolInvocation", ActivityKind.Internal);
        if (activity == null) return;

        activity.SetTag("tool.name", toolName);
        activity.SetTag("tool.action", action);
        activity.SetTag("tool.duration_ms", durationMs);
        activity.SetTag("tool.success", success);
        activity.SetTag("session.id", SessionId);
        activity.SetTag("app.version", GetVersion());

        activity.SetStatus(success ? ActivityStatusCode.Ok : ActivityStatusCode.Error);
    }

    /// <summary>
    /// Tracks an unhandled exception.
    /// Only call this for exceptions that escape all catch blocks (true bugs/crashes).
    /// </summary>
    /// <param name="exception">The unhandled exception</param>
    /// <param name="source">Source of the exception (e.g., "AppDomain.UnhandledException")</param>
    public static void TrackUnhandledException(Exception exception, string source)
    {
        if (!IsEnabled || exception == null) return;

        using var activity = ActivitySource.StartActivity("UnhandledException", ActivityKind.Internal);
        if (activity == null) return;

        // Redact sensitive data from exception
        var (type, message, stackTrace) = SensitiveDataRedactingProcessor.RedactException(exception);

        activity.SetTag("exception.type", type);
        activity.SetTag("exception.message", message);
        activity.SetTag("exception.source", source);
        activity.SetTag("session.id", SessionId);
        activity.SetTag("app.version", GetVersion());

        if (stackTrace != null)
        {
            // Truncate stack trace to avoid exceeding limits
            const int maxStackTraceLength = 4096;
            if (stackTrace.Length > maxStackTraceLength)
            {
                stackTrace = stackTrace[..maxStackTraceLength] + "... [truncated]";
            }
            activity.SetTag("exception.stacktrace", stackTrace);
        }

        activity.SetStatus(ActivityStatusCode.Error, message);

        // Record as exception event
        activity.AddEvent(new ActivityEvent("exception", tags: new ActivityTagsCollection
        {
            { "exception.type", type },
            { "exception.message", message }
        }));
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
}
