using System.Diagnostics.CodeAnalysis;

namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class WorksheetCommandValidation
{
    internal static string NamingRejectionMessage(string name, string existingName, bool isNewSheet) =>
        isNewSheet
            ? $"Excel rejected worksheet name '{name}'. The new worksheet '{existingName}' remains in the workbook; inspect it and remove it if appropriate."
            : $"Excel rejected worksheet name '{name}'. Worksheet '{existingName}' was not renamed.";

    internal static void ValidateNewSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name) ||
            name.Length > 31 ||
            name.AsSpan().IndexOfAny(":\\/?*[]".AsSpan()) >= 0 ||
            name.StartsWith('\'') ||
            name.EndsWith('\'') ||
            string.Equals(name, "History", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                "Worksheet names must be nonblank, contain at most 31 characters, not be 'History', " +
                "not contain : \\ / ? * [ ], and not begin or end with an apostrophe.",
                nameof(name));
        }
    }

    internal static void RequireAvailableName(string existingName, string name, string? currentName = null)
    {
        if (string.Equals(existingName, name, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(existingName, currentName, StringComparison.Ordinal))
            throw new InvalidOperationException($"Sheet '{name}' already exists.");
    }

    internal static void RequireExistingSheet([DoesNotReturnIf(false)] bool exists, string sheetName)
    {
        if (!exists)
        {
            throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
        }
    }
}
