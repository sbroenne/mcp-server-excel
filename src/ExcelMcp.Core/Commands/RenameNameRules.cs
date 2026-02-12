namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Shared helpers for rename validation and normalization rules.
/// </summary>
public static class RenameNameRules
{
    /// <summary>
    /// Trim leading/trailing whitespace; returns empty string for null input.
    /// </summary>
    public static string Normalize(string? name)
    {
        return name?.Trim() ?? string.Empty;
    }

    /// <summary>
    /// True when the normalized value is empty.
    /// </summary>
    public static bool IsEmpty(string normalizedName)
    {
        return string.IsNullOrWhiteSpace(normalizedName);
    }

    /// <summary>
    /// Determines if rename is a no-op (trimmed names identical including casing).
    /// </summary>
    public static bool IsNoOp(string normalizedOldName, string normalizedNewName)
    {
        return string.Equals(normalizedOldName, normalizedNewName, StringComparison.Ordinal);
    }

    /// <summary>
    /// Checks for case-insensitive conflicts after trimming, excluding the target being renamed.
    /// </summary>
    public static bool HasConflict(IEnumerable<string> existingNames, string normalizedNewName, string normalizedTargetName)
    {
        return existingNames
            .Select(Normalize)
            .Where(normalizedExisting => !string.Equals(normalizedExisting, normalizedTargetName, StringComparison.OrdinalIgnoreCase))
            .Any(normalizedExisting => string.Equals(normalizedExisting, normalizedNewName, StringComparison.OrdinalIgnoreCase));
    }
}


