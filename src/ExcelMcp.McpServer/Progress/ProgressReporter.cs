using System.Text.Json;

namespace Sbroenne.ExcelMcp.McpServer.Progress;

/// <summary>
/// Progress reporting helper for MCP server operations.
/// Provides standardized progress updates for long-running operations.
/// </summary>
public static class ProgressReporter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    /// <summary>
    /// Report progress for batch operations.
    /// </summary>
    public static void ReportBatchProgress(int current, int total, string? message = null)
    {
        var progress = new
        {
            current,
            total,
            percentage = total > 0 ? (double)current / total * 100 : 0,
            message = message ?? $"Processing {current} of {total}..."
        };

        // Write to stderr so it doesn't interfere with stdio MCP protocol
        Console.Error.WriteLine($"[PROGRESS] {JsonSerializer.Serialize(progress, JsonOptions)}");
    }

    /// <summary>
    /// Report operation start.
    /// </summary>
    public static void ReportStart(string operation, string? details = null)
    {
        var info = new
        {
            status = "started",
            operation,
            details,
            timestamp = DateTime.UtcNow
        };

        Console.Error.WriteLine($"[START] {JsonSerializer.Serialize(info, JsonOptions)}");
    }

    /// <summary>
    /// Report operation completion.
    /// </summary>
    public static void ReportComplete(string operation, TimeSpan duration, string? details = null)
    {
        var info = new
        {
            status = "completed",
            operation,
            durationMs = duration.TotalMilliseconds,
            details,
            timestamp = DateTime.UtcNow
        };

        Console.Error.WriteLine($"[COMPLETE] {JsonSerializer.Serialize(info, JsonOptions)}");
    }

    /// <summary>
    /// Report operation error.
    /// </summary>
    public static void ReportError(string operation, string error)
    {
        var info = new
        {
            status = "error",
            operation,
            error,
            timestamp = DateTime.UtcNow
        };

        Console.Error.WriteLine($"[ERROR] {JsonSerializer.Serialize(info, JsonOptions)}");
    }
}
