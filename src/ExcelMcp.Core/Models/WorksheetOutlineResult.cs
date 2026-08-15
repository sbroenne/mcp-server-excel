namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Worksheet outline state for a row or column range.
/// </summary>
public class WorksheetOutlineResult : OperationResult
{
    /// <summary>Name of the worksheet.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Row or column range inspected.</summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>Outline axis inspected.</summary>
    public OutlineAxis Axis { get; set; }

    /// <summary>Current outline level. Level 1 means the range is not grouped.</summary>
    public int OutlineLevel { get; set; }

    /// <summary>Whether the inspected rows or columns are hidden.</summary>
    public bool Hidden { get; set; }

    /// <summary>Summary row position: above or below.</summary>
    public string SummaryRow { get; set; } = string.Empty;

    /// <summary>Summary column position: left or right.</summary>
    public string SummaryColumn { get; set; } = string.Empty;

    /// <summary>Whether Excel applies automatic outline styles.</summary>
    public bool AutomaticStyles { get; set; }
}
