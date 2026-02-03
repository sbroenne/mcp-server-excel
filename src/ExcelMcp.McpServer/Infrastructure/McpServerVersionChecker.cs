using System.Reflection;

namespace Sbroenne.ExcelMcp.McpServer.Infrastructure;

/// <summary>
/// Checks for MCP Server updates and provides version information.
/// </summary>
public static class McpServerVersionChecker
{
    /// <summary>
    /// Checks for updates and returns update information if available.
    /// This method is non-blocking and fails silently if the check cannot be completed.
    /// </summary>
    /// <returns>Update info if update is available, null otherwise.</returns>
    public static async Task<UpdateInfo?> CheckForUpdateAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            var currentVersion = GetCurrentVersion();
            var latestVersion = await NuGetVersionChecker.GetLatestVersionAsync(cancellationToken);

            if (latestVersion == null)
            {
                // Could not check (network error, timeout, etc.)
                return null;
            }

            if (CompareVersions(currentVersion, latestVersion) < 0)
            {
                // Update available
                return new UpdateInfo
                {
                    CurrentVersion = currentVersion,
                    LatestVersion = latestVersion,
                    UpdateAvailable = true
                };
            }

            // Already up to date
            return null;
        }
        catch
        {
            // Fail silently - version check should never block server operation
            return null;
        }
    }

    /// <summary>
    /// Gets the current version of the MCP Server.
    /// </summary>
    public static string GetCurrentVersion()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        // Strip git hash suffix (e.g., "1.2.0+abc123" -> "1.2.0")
        return informational?.Split('+')[0] ?? assembly.GetName().Version?.ToString() ?? "0.0.0";
    }

    private static int CompareVersions(string current, string latest)
    {
        if (Version.TryParse(current, out var currentVer) && Version.TryParse(latest, out var latestVer))
            return currentVer.CompareTo(latestVer);
        return string.Compare(current, latest, StringComparison.Ordinal);
    }
}

/// <summary>
/// Information about an available update.
/// </summary>
public sealed class UpdateInfo
{
    public required string CurrentVersion { get; init; }
    public required string LatestVersion { get; init; }
    public required bool UpdateAvailable { get; init; }

    /// <summary>
    /// Gets a formatted message for logging or display.
    /// </summary>
    public string GetUpdateMessage() =>
        $"Update available: {CurrentVersion} -> {LatestVersion}. " +
        "Run: dotnet tool update --global Sbroenne.ExcelMcp.McpServer";
}
