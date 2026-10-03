using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Window;

/// <summary>Context from the session workbook's own windows, without activation.</summary>
public sealed class WindowContextResult : OperationResult
{
    /// <summary>available or no-windows.</summary>
    public string Availability { get; set; } = string.Empty;
    /// <summary>Visibility of the session's Excel application.</summary>
    public bool IsApplicationVisible { get; set; }
    /// <summary>Every window belonging to the session workbook.</summary>
    public List<WorkbookWindowContext> Windows { get; set; } = [];
}

/// <summary>Native context of one owned workbook window.</summary>
public sealed class WorkbookWindowContext
{
    /// <summary>Native workbook window number.</summary>
    public int WindowNumber { get; set; }
    /// <summary>Whether the workbook window is visible.</summary>
    public bool IsVisible { get; set; }
    /// <summary>worksheet, chart-sheet, or unsupported.</summary>
    public string SheetKind { get; set; } = string.Empty;
    /// <summary>Native active sheet name where supported.</summary>
    public string? SheetName { get; set; }
    /// <summary>range, chart, shapes, unsupported, or unavailable.</summary>
    public string SelectionKind { get; set; } = string.Empty;
    /// <summary>Native address only when the actual selection is a range.</summary>
    public string? SelectedRangeAddress { get; set; }
    /// <summary>Native active cell when the actual selection is a range.</summary>
    public string? ActiveCellAddress { get; set; }
    /// <summary>Native active chart name, if any.</summary>
    public string? ActiveChartName { get; set; }
    /// <summary>Why a native selection cannot be reported.</summary>
    public string? SelectionDiagnostic { get; set; }
}
