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

        return CredentialPattern().Replace(connectionString, "$1=(redacted)");
    }

    [GeneratedRegex(@"\b(password|pwd|user\s+id|uid)\s*=\s*(?:""(?:[^""]|"""")*""|'(?:[^']|'')*'|\{(?:[^}]|}})*\}|[^;]*)", RegexOptions.IgnoreCase)]
    private static partial Regex CredentialPattern();
}
