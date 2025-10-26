using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands
/// </summary>
public interface IParameterCommands
{
    /// <summary>
    /// Lists all named ranges in the workbook
    /// </summary>
    Task<ParameterListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Sets the value of a named range
    /// </summary>
    Task<OperationResult> SetAsync(IExcelBatch batch, string paramName, string value);

    /// <summary>
    /// Gets the value of a named range
    /// </summary>
    Task<ParameterValueResult> GetAsync(IExcelBatch batch, string paramName);

    /// <summary>
    /// Updates a named range reference
    /// </summary>
    Task<OperationResult> UpdateAsync(IExcelBatch batch, string paramName, string reference);

    /// <summary>
    /// Creates a new named range
    /// </summary>
    Task<OperationResult> CreateAsync(IExcelBatch batch, string paramName, string reference);

    /// <summary>
    /// Deletes a named range
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string paramName);
}
