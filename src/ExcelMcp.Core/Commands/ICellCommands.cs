using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Individual cell operation commands
/// </summary>
public interface ICellCommands
{
    /// <summary>
    /// Gets the value of a specific cell
    /// </summary>
    Task<CellValueResult> GetValueAsync(IExcelBatch batch, string sheetName, string cellAddress);

    /// <summary>
    /// Sets the value of a specific cell
    /// </summary>
    Task<OperationResult> SetValueAsync(IExcelBatch batch, string sheetName, string cellAddress, string value);

    /// <summary>
    /// Gets the formula of a specific cell
    /// </summary>
    Task<CellValueResult> GetFormulaAsync(IExcelBatch batch, string sheetName, string cellAddress);

    /// <summary>
    /// Sets the formula of a specific cell
    /// </summary>
    Task<OperationResult> SetFormulaAsync(IExcelBatch batch, string sheetName, string cellAddress, string formula);
}
