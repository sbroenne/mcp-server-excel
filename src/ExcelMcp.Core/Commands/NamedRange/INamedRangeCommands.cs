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
    NamedRangeListResult List(IExcelBatch batch);

    /// <summary>
    /// Sets the value of a named range
    /// </summary>
    OperationResult Write(IExcelBatch batch, string paramName, string value);

    /// <summary>
    /// Gets the value of a named range
    /// </summary>
    NamedRangeValueResult Read(IExcelBatch batch, string paramName);

    /// <summary>
    /// Updates a named range reference
    /// </summary>
    OperationResult Update(IExcelBatch batch, string paramName, string reference);

    /// <summary>
    /// Creates a new named range
    /// </summary>
    OperationResult Create(IExcelBatch batch, string paramName, string reference);

    /// <summary>
    /// Deletes a named range
    /// </summary>
    OperationResult Delete(IExcelBatch batch, string paramName);
}

