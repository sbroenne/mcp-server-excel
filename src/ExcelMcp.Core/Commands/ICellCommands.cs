using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Individual cell operation commands
/// </summary>
public interface ICellCommands
{
    /// <summary>
    /// Gets the value of a specific cell
    /// </summary>
    CellValueResult GetValue(string filePath, string sheetName, string cellAddress);

    /// <summary>
    /// Sets the value of a specific cell
    /// </summary>
    OperationResult SetValue(string filePath, string sheetName, string cellAddress, string value);

    /// <summary>
    /// Gets the formula of a specific cell
    /// </summary>
    CellValueResult GetFormula(string filePath, string sheetName, string cellAddress);

    /// <summary>
    /// Sets the formula of a specific cell
    /// </summary>
    OperationResult SetFormula(string filePath, string sheetName, string cellAddress, string formula);

    /// <summary>
    /// Sets the background color of a cell or range
    /// </summary>
    OperationResult SetBackgroundColor(string filePath, string sheetName, string cellAddress, string color);

    /// <summary>
    /// Sets the font color of a cell or range
    /// </summary>
    OperationResult SetFontColor(string filePath, string sheetName, string cellAddress, string color);

    /// <summary>
    /// Sets the font properties of a cell or range
    /// </summary>
    OperationResult SetFont(string filePath, string sheetName, string cellAddress, string? fontName = null, int? fontSize = null, bool? bold = null, bool? italic = null, bool? underline = null);

    /// <summary>
    /// Sets the border style for a cell or range
    /// </summary>
    OperationResult SetBorder(string filePath, string sheetName, string cellAddress, string borderStyle, string? borderColor = null);

    /// <summary>
    /// Sets the number format for a cell or range
    /// </summary>
    OperationResult SetNumberFormat(string filePath, string sheetName, string cellAddress, string format);

    /// <summary>
    /// Sets the alignment for a cell or range
    /// </summary>
    OperationResult SetAlignment(string filePath, string sheetName, string cellAddress, string? horizontal = null, string? vertical = null, bool? wrapText = null);

    /// <summary>
    /// Clears all formatting from a cell or range
    /// </summary>
    OperationResult ClearFormatting(string filePath, string sheetName, string cellAddress);
}
