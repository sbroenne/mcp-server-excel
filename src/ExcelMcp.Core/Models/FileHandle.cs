namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Represents an open Excel workbook file handle.
/// Internal identifier for tracking workbook lifecycle in the active workbook pattern.
/// </summary>
public sealed class FileHandle
{
    /// <summary>
    /// Unique identifier for this file handle
    /// </summary>
    public string Id { get; }

    /// <summary>
    /// Absolute path to the Excel workbook file
    /// </summary>
    public string FilePath { get; }

    /// <summary>
    /// UTC timestamp when the file was opened
    /// </summary>
    public DateTime OpenedAt { get; }

    /// <summary>
    /// Indicates whether this handle has been closed
    /// </summary>
    public bool IsClosed { get; internal set; }

    /// <summary>
    /// Creates a new file handle
    /// </summary>
    /// <param name="id">Unique identifier</param>
    /// <param name="filePath">Path to the workbook file</param>
    internal FileHandle(string id, string filePath)
    {
        Id = id ?? throw new ArgumentNullException(nameof(id));
        FilePath = Path.GetFullPath(filePath);
        OpenedAt = DateTime.UtcNow;
        IsClosed = false;
    }
}
