namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Worksheet-specific view state for an Excel workbook window.
/// </summary>
public class WorksheetViewResult : OperationResult
{
    /// <summary>Name of the active worksheet whose view was inspected.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Whether panes are frozen.</summary>
    public bool FreezePanes { get; set; }

    /// <summary>Number of rows above the pane boundary.</summary>
    public int SplitRow { get; set; }

    /// <summary>Number of columns left of the pane boundary.</summary>
    public int SplitColumn { get; set; }

    /// <summary>Worksheet zoom percentage.</summary>
    public int Zoom { get; set; }

    /// <summary>Whether worksheet gridlines are displayed.</summary>
    public bool DisplayGridlines { get; set; }

    /// <summary>Whether row and column headings are displayed.</summary>
    public bool DisplayHeadings { get; set; }

    /// <summary>Whether outline symbols are displayed.</summary>
    public bool DisplayOutlineSymbols { get; set; }

    /// <summary>Whether formulas are displayed instead of their calculated values.</summary>
    public bool DisplayFormulas { get; set; }
}
