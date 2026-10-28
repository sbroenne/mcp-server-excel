namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Normalizes workbook identities without treating remote locations as disk paths.
/// </summary>
internal static class WorkbookLocation
{
    internal static bool IsRemote(string location) =>
        Uri.TryCreate(location, UriKind.Absolute, out var uri) && !uri.IsFile;

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
