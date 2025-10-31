using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model TOM (Tabular Object Model) commands - Calculated column operations
/// Provides calculated column support using Microsoft Analysis Services TOM API
/// Note: Measure and relationship operations use COM API (see DataModelCommands)
/// </summary>
public interface IDataModelTomCommands
{
    /// <summary>
    /// Creates a calculated column in a table using DAX
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="tableName">Table to add the column to</param>
    /// <param name="columnName">Name of the new calculated column</param>
    /// <param name="daxFormula">DAX expression for the column</param>
    /// <param name="description">Optional description for the column</param>
    /// <param name="dataType">Data type: String, Integer, Double, Boolean, DateTime</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult CreateCalculatedColumn(
        string filePath,
        string tableName,
        string columnName,
        string daxFormula,
        string? description = null,
        string dataType = "String");

    /// <summary>
    /// Lists all calculated columns in the Data Model
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="tableName">Optional table name to filter columns (null for all tables)</param>
    /// <returns>Result with list of calculated columns</returns>
    DataModelCalculatedColumnListResult ListCalculatedColumns(string filePath, string? tableName = null);

    /// <summary>
    /// Views details of a specific calculated column
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="tableName">Table containing the column</param>
    /// <param name="columnName">Name of the calculated column</param>
    /// <returns>Result with column details</returns>
    DataModelCalculatedColumnViewResult ViewCalculatedColumn(string filePath, string tableName, string columnName);

    /// <summary>
    /// Updates an existing calculated column's formula and properties
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="tableName">Table containing the column</param>
    /// <param name="columnName">Name of the calculated column to update</param>
    /// <param name="daxFormula">New DAX expression (null to keep existing)</param>
    /// <param name="description">New description (null to keep existing)</param>
    /// <param name="dataType">New data type (null to keep existing)</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult UpdateCalculatedColumn(
        string filePath,
        string tableName,
        string columnName,
        string? daxFormula = null,
        string? description = null,
        string? dataType = null);

    /// <summary>
    /// Deletes a calculated column from a table
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="tableName">Table containing the column</param>
    /// <param name="columnName">Name of the calculated column to delete</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult DeleteCalculatedColumn(string filePath, string tableName, string columnName);

    /// <summary>
    /// Validates a DAX formula without creating/updating any objects
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="daxFormula">DAX formula to validate</param>
    /// <returns>Result with validation status and any errors</returns>
    DataModelValidationResult ValidateDax(string filePath, string daxFormula);
}
