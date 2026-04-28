using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.Core.DataModel;

/// <summary>
/// Diagnostic information from Excel's Data Model ADO connection.
/// </summary>
internal sealed class DataModelAdoDiagnostics
{
    /// <summary>
    /// The MSOLAP provider name parsed from the ADO connection string, if available.
    /// </summary>
    public string? ProviderName { get; set; }

    /// <summary>
    /// The ADO connection string, if available. Sanitize before including it in output.
    /// </summary>
    public string? ConnectionString { get; set; }

    /// <summary>
    /// Parses the OLE DB provider name from a semicolon-separated connection string.
    /// </summary>
    internal static string? ExtractProviderName(string? connectionString)
    {
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            return null;
        }

        var match = Regex.Match(connectionString, @"(?:^|;)\s*Provider\s*=\s*([^;]+)", RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value.Trim().Trim('"') : null;
    }

    /// <summary>
    /// Removes credentials from a connection string before it is included in error output.
    /// </summary>
    internal static string? SanitizeConnectionString(string? connectionString)
    {
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            return connectionString;
        }

        return Regex.Replace(
            connectionString,
            @"(?i)\b(password|pwd|user\s+id|uid)\s*=\s*[^;]*",
            "$1=(redacted)");
    }
}
