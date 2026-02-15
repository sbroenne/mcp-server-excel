namespace Sbroenne.ExcelMcp.Service.Infrastructure;

/// <summary>
/// Information about an available update.
/// </summary>
public sealed class UpdateInfo
{
    public required string CurrentVersion { get; init; }
    public required string LatestVersion { get; init; }
    public required bool UpdateAvailable { get; init; }

    /// <summary>
    /// Gets a notification title for Windows toast notifications.
    /// </summary>
    public static string GetNotificationTitle() => "Excel MCP Update Available";

    /// <summary>
    /// Gets a formatted message for logging or display (generic version).
    /// </summary>
    public string GetUpdateMessage() =>
        $"Update available: {CurrentVersion} -> {LatestVersion}";

    /// <summary>
    /// Gets a formatted message for logging or display with specific update command.
    /// </summary>
    public string GetUpdateMessage(string updateCommand) =>
        $"Update available: {CurrentVersion} -> {LatestVersion}. Run: {updateCommand}";

    /// <summary>
    /// Gets a formatted notification message for Windows toast notifications.
    /// </summary>
    public string GetNotificationMessage() =>
        $"Version {LatestVersion} is available (current: {CurrentVersion}).\n" +
        "Update both packages:\n" +
        "dotnet tool update --global Sbroenne.ExcelMcp.McpServer\n" +
        "dotnet tool update --global Sbroenne.ExcelMcp.CLI";
}


