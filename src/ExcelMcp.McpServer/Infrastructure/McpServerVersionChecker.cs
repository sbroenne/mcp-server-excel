using System.Reflection;

namespace Sbroenne.ExcelMcp.McpServer.Infrastructure;

/// <summary>
/// Checks for MCP Server updates and provides version information.
/// </summary>
public static class McpServerVersionChecker
{
    /// <summary>
    /// Checks for updates and returns the latest version if an update is available.
    /// </summary>
    /// <returns>Latest version string if update available, null otherwise.</returns>
    public static async Task<string?> CheckForUpdateAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            var currentVersion = GetCurrentVersion();
            var latestVersion = await NuGetVersionChecker.GetLatestVersionAsync(cancellationToken);

            if (latestVersion != null && CompareVersions(currentVersion, latestVersion) < 0)
            {
                return latestVersion;
            }

            return null;
        }
        catch (Exception)
        {
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
