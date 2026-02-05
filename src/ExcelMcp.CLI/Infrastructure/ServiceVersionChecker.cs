using System.Reflection;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Checks for CLI updates on service startup and notifies via Windows notification.
/// </summary>
internal static class ServiceVersionChecker
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
            // Fail silently - version check should never block service startup
            return null;
        }
    }

    private static string GetCurrentVersion()
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
internal sealed class UpdateInfo
{
    public required string CurrentVersion { get; init; }
    public required string LatestVersion { get; init; }
    public required bool UpdateAvailable { get; init; }

    public static string GetNotificationTitle() => "Excel CLI Update Available";

    public string GetNotificationMessage() =>
        $"Version {LatestVersion} is available (current: {CurrentVersion}).\n" +
        "Update via: dotnet tool update --global Sbroenne.ExcelMcp.CLI";
}
