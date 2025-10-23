using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands
/// </summary>
public interface IParameterCommands
{
    /// <summary>
    /// Lists all named ranges in the workbook
    /// </summary>
    ParameterListResult List(string filePath);

    /// <summary>
    /// Sets the value of a named range
    /// </summary>
    OperationResult Set(string filePath, string paramName, string value);

    /// <summary>
    /// Gets the value of a named range
    /// </summary>
    ParameterValueResult Get(string filePath, string paramName);

    /// <summary>
    /// Updates a named range reference
    /// </summary>
    OperationResult Update(string filePath, string paramName, string reference);

    /// <summary>
    /// Creates a new named range
    /// </summary>
    OperationResult Create(string filePath, string paramName, string reference);

    /// <summary>
    /// Deletes a named range
    /// </summary>
    OperationResult Delete(string filePath, string paramName);
}
