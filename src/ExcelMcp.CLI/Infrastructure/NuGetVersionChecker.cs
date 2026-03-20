using System.Net.Http.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Checks GitHub Releases for the latest version of the CLI.
/// </summary>
internal static class NuGetVersionChecker
{
    private const string LatestReleaseUrl = "https://api.github.com/repos/sbroenne/mcp-server-excel/releases/latest";
    private const string UserAgent = "ExcelMcp-CLI";
    private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Checks GitHub Releases for the latest version.
    /// </summary>
    /// <returns>Latest version string, or null if check failed.</returns>
    public static async Task<string?> GetLatestVersionAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            using var httpClient = new HttpClient { Timeout = Timeout };
            httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(UserAgent);

            var response = await httpClient.GetFromJsonAsync<GitHubReleaseResponse>(LatestReleaseUrl, cancellationToken);

            if (response?.TagName == null)
                return null;

            // Strip 'v' prefix: "v1.2.3" -> "1.2.3"
            return response.TagName.TrimStart('v');
        }
        catch (Exception)
        {
            // Network error, timeout, etc. — return null to indicate check failed
            return null;
        }
    }

    private sealed class GitHubReleaseResponse
    {
        [JsonPropertyName("tag_name")]
        public string? TagName { get; set; }
    }
}
