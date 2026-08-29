namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Expected numeric value for one control-total cell.</summary>
public sealed class WorkbookControlTotalExpectation
{
    /// <summary>Worksheet containing the control-total cell.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Single-cell A1 address to compare.</summary>
    public string CellAddress { get; set; } = string.Empty;

    /// <summary>Required expected finite numeric value.</summary>
    public double? ExpectedValue { get; set; }

    /// <summary>Allowed non-negative absolute difference from the expected value.</summary>
    public double Tolerance { get; set; }
}
