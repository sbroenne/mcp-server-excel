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
}
