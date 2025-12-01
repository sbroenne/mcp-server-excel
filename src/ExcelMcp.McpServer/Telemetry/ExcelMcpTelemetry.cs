// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Reflection;
using Microsoft.ApplicationInsights;

namespace Sbroenne.ExcelMcp.McpServer.Telemetry;

/// <summary>
/// Centralized telemetry helper for ExcelMcp MCP Server.
/// Provides usage tracking and unhandled exception reporting via Application Insights SDK.
/// </summary>
public static class ExcelMcpTelemetry
{
    /// <summary>
    /// Environment variable to enable debug mode (console output instead of Azure).
    /// Set to "true" or "1" for local testing without Azure resources.
    /// </summary>
    public const string DebugTelemetryEnvironmentVariable = "EXCELMCP_DEBUG_TELEMETRY";

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
    /// Called by Program.cs during startup.
    /// </summary>
    internal static void SetTelemetryClient(TelemetryClient client)
    {
        _telemetryClient = client;
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
            if (IsDebugMode())
            {
                Console.Error.WriteLine("[Telemetry] Flushing telemetry buffer...");
            }

            // Flush with timeout to avoid hanging on shutdown
            // 5 seconds is typically sufficient for small batches
            _telemetryClient.FlushAsync(CancellationToken.None).Wait(TimeSpan.FromSeconds(5));

            if (IsDebugMode())
            {
                Console.Error.WriteLine("[Telemetry] Telemetry buffer flushed successfully");
            }
        }
        catch (Exception ex)
        {
            // Don't let telemetry flush failure crash the application
            if (IsDebugMode())
            {
                Console.Error.WriteLine($"[Telemetry] Warning: Failed to flush telemetry: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Gets whether telemetry is enabled. Always returns true.
    /// </summary>
    public static bool IsEnabled => true;

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
        // Connection string is embedded at build time from APPINSIGHTS_CONNECTION_STRING env var
        // Returns null if not set during build (local dev builds)
        return string.IsNullOrEmpty(TelemetryConfig.ConnectionString) ? null : TelemetryConfig.ConnectionString;
    }

    /// <summary>
    /// Tracks a tool invocation with usage metrics.
    /// Sends Application Insights Request and PageView telemetry.
    /// - Request: Populates Performance, Failures, Users, Sessions blades
    /// - PageView: Enables User Flows blade (shows tool usage patterns)
    /// </summary>
    /// <param name="toolName">The MCP tool name (e.g., "excel_range")</param>
    /// <param name="action">The action performed (e.g., "get-values")</param>
    /// <param name="durationMs">Duration in milliseconds</param>
    /// <param name="success">Whether the operation succeeded</param>
    public static void TrackToolInvocation(string toolName, string action, long durationMs, bool success)
    {
        if (!IsEnabled) return;

        // Debug mode: write to stderr
        if (IsDebugMode())
        {
            Console.Error.WriteLine($"[Telemetry] ToolInvocation: {toolName}.{action} - {(success ? "Success" : "Failed")} ({durationMs}ms)");
        }

        if (_telemetryClient != null)
        {
            var operationName = $"{toolName}/{action}";
            var startTime = DateTimeOffset.UtcNow.AddMilliseconds(-durationMs);
            var duration = TimeSpan.FromMilliseconds(durationMs);

            // Request telemetry: Performance, Failures, Users, Sessions
            _telemetryClient.TrackRequest(operationName, startTime, duration, success ? "200" : "500", success);

            // PageView telemetry: Enables User Flows blade
            _telemetryClient.TrackPageView(operationName);
        }
    }

    /// <summary>
    /// Tracks an unhandled exception.
    /// Only call this for exceptions that escape all catch blocks (true bugs/crashes).
    /// Sends Application Insights exception and Custom Event.
    /// </summary>
    /// <param name="exception">The unhandled exception</param>
    /// <param name="source">Source of the exception (e.g., "AppDomain.UnhandledException")</param>
    public static void TrackUnhandledException(Exception exception, string source)
    {
        if (!IsEnabled || exception == null) return;

        // Redact sensitive data from exception
        var (type, message, _) = SensitiveDataRedactor.RedactException(exception);

        // Debug mode: write to stderr
        if (IsDebugMode())
        {
            Console.Error.WriteLine($"[Telemetry] UnhandledException: {type} - {message} (Source: {source})");
        }

        // Send Application Insights telemetry
        if (_telemetryClient != null)
        {
            // Track as exception in Application Insights (for Failures blade)
            _telemetryClient.TrackException(exception, new Dictionary<string, string>
            {
                { "Source", source },
                { "ExceptionType", type },
                { "AppVersion", GetVersion() }
            });
        }
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
            var bytes = System.Text.Encoding.UTF8.GetBytes(machineIdentity);
            var hash = System.Security.Cryptography.SHA256.HashData(bytes);
            return Convert.ToHexString(hash)[..16].ToLowerInvariant();
        }
        catch
        {
            // Fallback to a random ID if machine identity cannot be determined
            return Guid.NewGuid().ToString("N")[..16];
        }
    }
}
