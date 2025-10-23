using System.Reflection;
using NuGet.Common;
using NuGet.Protocol;
using NuGet.Protocol.Core.Types;
using NuGet.Versioning;

namespace Sbroenne.ExcelMcp.Core;

/// <summary>
/// Checks for the latest version of the ExcelMcp package on NuGet.org
/// </summary>
public class VersionChecker
{
    private const string NuGetSource = "https://api.nuget.org/v3/index.json";
    private readonly ILogger _logger;

    /// <summary>
    /// Creates a new instance of the VersionChecker
    /// </summary>
    /// <param name="logger">Optional logger for diagnostic output</param>
    public VersionChecker(ILogger? logger = null)
    {
        _logger = logger ?? NullLogger.Instance;
    }

    /// <summary>
    /// Gets the current version of the running assembly
    /// </summary>
    public static Version GetCurrentVersion()
    {
        var assembly = Assembly.GetEntryAssembly() ?? Assembly.GetExecutingAssembly();
        var version = assembly.GetName().Version;
        return version ?? new Version(1, 0, 0, 0);
    }

    /// <summary>
    /// Checks if a newer version is available on NuGet.org
    /// </summary>
    /// <param name="packageId">The NuGet package ID to check</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Version check result</returns>
    public async Task<VersionCheckResult> CheckForUpdatesAsync(
        string packageId,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var currentVersion = GetCurrentVersion();
            var cache = new SourceCacheContext();
            var repository = Repository.Factory.GetCoreV3(NuGetSource);
            var resource = await repository.GetResourceAsync<FindPackageByIdResource>(cancellationToken);

            // Get all versions
            var versions = await resource.GetAllVersionsAsync(
                packageId,
                cache,
                _logger,
                cancellationToken);

            if (versions == null || !versions.Any())
            {
                return new VersionCheckResult
                {
                    Success = false,
                    ErrorMessage = $"Package '{packageId}' not found on NuGet.org"
                };
            }

            // Get the latest stable version (exclude pre-release)
            var latestVersion = versions
                .Where(v => !v.IsPrerelease)
                .OrderByDescending(v => v)
                .FirstOrDefault();

            if (latestVersion == null)
            {
                // If no stable version, get the latest pre-release
                latestVersion = versions.OrderByDescending(v => v).First();
            }

            var currentNuGetVersion = new NuGetVersion(currentVersion);
            var isOutdated = latestVersion > currentNuGetVersion;

            return new VersionCheckResult
            {
                Success = true,
                CurrentVersion = currentVersion.ToString(),
                LatestVersion = latestVersion.ToString(),
                IsOutdated = isOutdated,
                PackageId = packageId
            };
        }
        catch (Exception ex)
        {
            return new VersionCheckResult
            {
                Success = false,
                ErrorMessage = $"Failed to check for updates: {ex.Message}",
                CurrentVersion = GetCurrentVersion().ToString()
            };
        }
    }
}

/// <summary>
/// Result of a version check operation
/// </summary>
public class VersionCheckResult
{
    /// <summary>
    /// Whether the version check was successful
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// Current version running
    /// </summary>
    public string? CurrentVersion { get; set; }

    /// <summary>
    /// Latest version available on NuGet.org
    /// </summary>
    public string? LatestVersion { get; set; }

    /// <summary>
    /// Whether the current version is outdated
    /// </summary>
    public bool IsOutdated { get; set; }

    /// <summary>
    /// The package ID that was checked
    /// </summary>
    public string? PackageId { get; set; }

    /// <summary>
    /// Error message if the check failed
    /// </summary>
    public string? ErrorMessage { get; set; }

    /// <summary>
    /// Gets a user-friendly message about the version status
    /// </summary>
    public string GetMessage()
    {
        if (!Success)
        {
            return $"Warning: Could not check for updates. {ErrorMessage}";
        }

        if (IsOutdated)
        {
            return $"Warning: A newer version ({LatestVersion}) is available. You are running version {CurrentVersion}. " +
                   $"Update with: dotnet tool update -g {PackageId}";
        }

        return $"You are running the latest version ({CurrentVersion}).";
    }
}
