using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands
/// </summary>
public interface INamedRangeCommands
{
    /// <summary>
    /// Lists all named ranges in the workbook
    /// </summary>
    Task<NamedRangeListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Lists all named ranges in the workbook (filePath-based API)
    /// </summary>
    Task<NamedRangeListResult> ListAsync(string filePath);

    /// <summary>
    /// Sets the value of a named range
    /// </summary>
    Task<OperationResult> SetAsync(IExcelBatch batch, string paramName, string value);

    /// <summary>
    /// Sets the value of a named range (filePath-based API)
    /// </summary>
    Task<OperationResult> SetAsync(string filePath, string paramName, string value);

    /// <summary>
    /// Gets the value of a named range
    /// </summary>
    Task<NamedRangeValueResult> GetAsync(IExcelBatch batch, string paramName);

    /// <summary>
    /// Gets the value of a named range (filePath-based API)
    /// </summary>
    Task<NamedRangeValueResult> GetAsync(string filePath, string paramName);

    /// <summary>
    /// Updates a named range reference
    /// </summary>
    Task<OperationResult> UpdateAsync(IExcelBatch batch, string paramName, string reference);

    /// <summary>
    /// Updates a named range reference (filePath-based API)
    /// </summary>
    Task<OperationResult> UpdateAsync(string filePath, string paramName, string reference);

    /// <summary>
    /// Creates a new named range
    /// </summary>
    Task<OperationResult> CreateAsync(IExcelBatch batch, string paramName, string reference);

    /// <summary>
    /// Creates a new named range (filePath-based API)
    /// </summary>
    Task<OperationResult> CreateAsync(string filePath, string paramName, string reference);

    /// <summary>
    /// Deletes a named range
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string paramName);

    /// <summary>
    /// Deletes a named range (filePath-based API)
    /// </summary>
    Task<OperationResult> DeleteAsync(string filePath, string paramName);
}
