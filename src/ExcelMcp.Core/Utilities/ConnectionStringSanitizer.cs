using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>
/// Removes credentials from connection strings before they are returned to clients.
/// </summary>
internal static partial class ConnectionStringSanitizer
{
    internal static string? Sanitize(string? connectionString)
    {
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            return connectionString;
        }

        var sanitized = UriUserInfoPattern().Replace(connectionString, "$1(redacted)@");
        sanitized = UriQueryCredentialPattern().Replace(sanitized, "$1(redacted)");
        return CredentialPattern().Replace(sanitized, "${prefix}${key}=(redacted)");
    }

    [GeneratedRegex(@"(?<prefix>(?:^|;)\s*)(?<key>password|pwd|user\s+id|uid|user\s*name|username|user|client[\s_-]*secret|access[\s_-]*token|auth[\s_-]*token|bearer[\s_-]*token|api[\s_-]*key|account[\s_-]*key|subscription[\s_-]*key|shared[\s_-]*access[\s_-]*signature|signature|secret|token)\s*=\s*(?:""(?:[^""]|"""")*""|'(?:[^']|'')*'|\{(?:[^}]|}})*\}|[^;]*)", RegexOptions.IgnoreCase)]
    private static partial Regex CredentialPattern();

    [GeneratedRegex(@"(\b[a-z][a-z0-9+.-]*://)[^/@\s;]+@", RegexOptions.IgnoreCase)]
    private static partial Regex UriUserInfoPattern();

    [GeneratedRegex(@"([?&](?:password|pwd|user(?:[\s_-]*id|[\s_-]*name)?|uid|username|client[\s_-]*secret|access[\s_-]*token|auth[\s_-]*token|bearer[\s_-]*token|api[\s_-]*key|account[\s_-]*key|subscription[\s_-]*key|shared[\s_-]*access[\s_-]*signature|signature|secret|token|sig)=)[^&#;\s]*", RegexOptions.IgnoreCase)]
    private static partial Regex UriQueryCredentialPattern();
}
