using System.Net.Http.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Checks NuGet for the latest version of the CLI package.
/// </summary>
internal static class NuGetVersionChecker
{
    private const string PackageId = "sbroenne.excelmcp.cli";
    private const string NuGetIndexUrl = $"https://api.nuget.org/v3-flatcontainer/{PackageId}/index.json";
    private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Checks NuGet for the latest version of the CLI package.
    /// </summary>
    /// <returns>Latest version string, or null if check failed.</returns>
    public static async Task<string?> GetLatestVersionAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            using var httpClient = new HttpClient { Timeout = Timeout };
            var response = await httpClient.GetFromJsonAsync<NuGetVersionsResponse>(NuGetIndexUrl, cancellationToken);

            if (response?.Versions == null || response.Versions.Count == 0)
                return null;

            // Get highest non-prerelease version
            var latestVersion = response.Versions
                .Where(v => !v.Contains('-')) // Exclude prerelease versions
                .OrderByDescending(v => ParseVersion(v))
                .FirstOrDefault();

            return latestVersion ?? response.Versions.Last();
        }
        catch
        {
            // Network error, timeout, etc. - return null to indicate check failed
            return null;
        }
    }

    private static Version ParseVersion(string versionString)
    {
        // Handle versions like "1.2.3" - strip any suffix after +
        var cleanVersion = versionString.Split('+')[0].Split('-')[0];
        return Version.TryParse(cleanVersion, out var version) ? version : new Version(0, 0, 0);
    }

    private sealed class NuGetVersionsResponse
    {
        [JsonPropertyName("versions")]
        public List<string> Versions { get; set; } = [];
    }
}
