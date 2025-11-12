using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Session;

/// <summary>
/// Thread-safe active workbook tracking using AsyncLocal.
/// Each async call chain has its own isolated active workbook context.
/// </summary>
internal static class ActiveWorkbook
{
    private static readonly AsyncLocal<FileHandle?> _current = new();

    /// <summary>
    /// Gets or sets the current active workbook handle.
    /// Throws InvalidOperationException if no workbook is active when getting.
    /// </summary>
    internal static FileHandle Current
    {
        get => _current.Value ?? throw new InvalidOperationException(
            "No active workbook. Call OpenAsync() or CreateAsync() first.");
        set => _current.Value = value;
    }

    /// <summary>
    /// Checks if there is an active workbook in the current async context
    /// </summary>
    internal static bool HasActive => _current.Value != null;
}
