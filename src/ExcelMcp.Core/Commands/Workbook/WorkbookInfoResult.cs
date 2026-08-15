using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>
/// Active workbook metadata.
/// </summary>
public sealed class WorkbookInfoResult : ResultBase
{
    /// <summary>Workbook file name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Absolute workbook file path reported by Excel.</summary>
    public string FullName { get; set; } = string.Empty;

    /// <summary>Directory containing the workbook.</summary>
    public string DirectoryPath { get; set; } = string.Empty;

    /// <summary>Short workbook format name.</summary>
    public string Format { get; set; } = string.Empty;

    /// <summary>Excel XlFileFormat numeric value.</summary>
    public int FormatCode { get; set; }

    /// <summary>Whether Excel considers all workbook changes saved.</summary>
    public bool Saved { get; set; }

    /// <summary>Whether the workbook is open read-only.</summary>
    public bool ReadOnly { get; set; }

    /// <summary>Whether the workbook has an open password.</summary>
    public bool HasPassword { get; set; }

    /// <summary>Whether the workbook is write-reserved.</summary>
    public bool WriteReserved { get; set; }
}
