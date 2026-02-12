// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using Microsoft.ApplicationInsights;
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
/// User/Session context is set by ExcelMcpTelemetryInitializer on all telemetry items.
/// View data in Azure Portal: Logs blade with Kusto queries on customEvents/requests tables.
/// </summary>
public static class ExcelMcpTelemetry
{
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
    /// Uses TrackEvent directly - ITelemetryInitializer sets user/session context.
    /// This ensures Users/Sessions blades have data even if no tools are invoked.
    /// </summary>
    private static void TrackSessionStart()
    {
        if (_telemetryClient == null) return;

        _telemetryClient.TrackEvent("SessionStart", new Dictionary<string, string>
        {
            ["SessionId"] = SessionId,
            ["AppVersion"] = GetVersion()
        });
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
            // 5 seconds is typically sufficient for small batches
            _telemetryClient.FlushAsync(CancellationToken.None).Wait(TimeSpan.FromSeconds(5));
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

        var metrics = new Dictionary<string, double>
        {
            ["DurationMs"] = durationMs
        };

        // Track as customEvent for analytics (tool usage, parameters, success/failure)
        _telemetryClient.TrackEvent(operationName, properties, metrics);

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

        // Redact sensitive data from exception
        var (type, _, _) = SensitiveDataRedactor.RedactException(exception);

        // Track as exception in Application Insights (for Failures blade)
        _telemetryClient.TrackException(exception, new Dictionary<string, string>
        {
            ["Source"] = source,
            ["ExceptionType"] = type,
            ["AppVersion"] = GetVersion()
        });
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
}


