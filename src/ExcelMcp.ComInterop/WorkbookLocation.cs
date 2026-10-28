using System.Security.Cryptography;
using System.Text;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Separates reported workbook locations from private, query-sensitive ownership keys.
/// </summary>
internal static class WorkbookLocation
{
    internal static bool IsRemote(string location) =>
        Uri.TryCreate(location, UriKind.Absolute, out var uri) && !uri.IsFile;

    internal static string GetOwnershipKey(string location)
    {
        var normalized = Normalize(location);
        if (!IsRemote(location) || new Uri(location).Query is not { Length: > 0 } query)
            return normalized;

        return normalized + "#" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(query)));
    }

    internal static bool MatchesReportedLocation(string ownershipKey, string? location) =>
        string.Equals(Normalize(ownershipKey), location, StringComparison.OrdinalIgnoreCase);

    internal static string Normalize(string location)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(location);
        if (!Uri.TryCreate(location, UriKind.Absolute, out var uri) || uri.IsFile)
        {
            if (uri?.IsFile == true && !Path.IsPathFullyQualified(location))
                location = uri.LocalPath;
            if (!Path.IsPathFullyQualified(location))
                throw new ArgumentException("Workbook location must be an absolute Windows path or a URL.", nameof(location));
            return Path.GetFullPath(location);
        }

        return new UriBuilder(uri)
        {
            Path = string.Join('/', uri.AbsolutePath.Split('/')
                .Select(segment => Uri.EscapeDataString(Uri.UnescapeDataString(segment)))),
            UserName = string.Empty,
            Password = string.Empty,
            Query = string.Empty,
            Fragment = string.Empty
        }.Uri.AbsoluteUri;
    }

    internal static string GetExtension(string location) =>
        Path.GetExtension(IsRemote(location)
            ? Uri.UnescapeDataString(new Uri(location).AbsolutePath)
            : location).ToLowerInvariant();

}
