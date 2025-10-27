using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Excel range operation commands - bulk operations on cell ranges for performance
/// </summary>
public interface IRangeCommands
{
    /// <summary>
    /// Gets the values of all cells in a range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Range address (e.g., "A1:D10" or "A1" for single cell)</param>
    /// <returns>RangeValueResult containing 2D array of cell values</returns>
    Task<RangeValueResult> GetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Sets the values of all cells in a range from a 2D array
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Range address (e.g., "A1:D10")</param>
    /// <param name="values">2D array of values (row-major order). Can be jagged for partial ranges.</param>
    /// <returns>OperationResult indicating success/failure</returns>
    Task<OperationResult> SetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<object?>> values);

    /// <summary>
    /// Gets the formulas of all cells in a range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Range address (e.g., "A1:D10")</param>
    /// <returns>RangeFormulaResult containing formulas and calculated values</returns>
    Task<RangeFormulaResult> GetFormulasAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Sets the formulas of all cells in a range from a 2D array
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Range address (e.g., "A1:D10")</param>
    /// <param name="formulas">2D array of formulas (row-major order). Formulas should include '=' prefix.</param>
    /// <returns>OperationResult indicating success/failure</returns>
    Task<OperationResult> SetFormulasAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>> formulas);

    /// <summary>
    /// Clears all content (values and formulas) from a range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Range address (e.g., "A1:D10")</param>
    /// <returns>OperationResult indicating success/failure</returns>
    Task<OperationResult> ClearAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Copies a range to another location
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sourceSheet">Source worksheet name</param>
    /// <param name="sourceRange">Source range address</param>
    /// <param name="targetSheet">Target worksheet name</param>
    /// <param name="targetRange">Target range address (can be single cell for top-left)</param>
    /// <returns>OperationResult indicating success/failure</returns>
    Task<OperationResult> CopyAsync(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange);
}
