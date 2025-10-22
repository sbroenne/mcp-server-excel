using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model TOM (Tabular Object Model) commands - Advanced CRUD operations
/// Provides create and update capabilities using Microsoft Analysis Services TOM API
/// </summary>
public interface IDataModelTomCommands
{
    /// <summary>
    /// Creates a new DAX measure in the specified table
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="tableName">Table to add the measure to</param>
    /// <param name="measureName">Name of the new measure</param>
    /// <param name="daxFormula">DAX expression for the measure</param>
    /// <param name="description">Optional description for the measure</param>
    /// <param name="formatString">Optional format string (e.g., "#,##0.00", "0.0%")</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult CreateMeasure(
        string filePath,
        string tableName,
        string measureName,
        string daxFormula,
        string? description = null,
        string? formatString = null);

    /// <summary>
    /// Updates an existing DAX measure's formula and properties
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="measureName">Name of the measure to update</param>
    /// <param name="daxFormula">New DAX expression (null to keep existing)</param>
    /// <param name="description">New description (null to keep existing)</param>
    /// <param name="formatString">New format string (null to keep existing)</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult UpdateMeasure(
        string filePath,
        string measureName,
        string? daxFormula = null,
        string? description = null,
        string? formatString = null);

    /// <summary>
    /// Creates a new relationship between two tables
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="fromTable">Source table name (many side)</param>
    /// <param name="fromColumn">Source column name (foreign key)</param>
    /// <param name="toTable">Target table name (one side)</param>
    /// <param name="toColumn">Target column name (primary key)</param>
    /// <param name="isActive">Whether relationship is active (default: true)</param>
    /// <param name="crossFilterDirection">Filter direction: Single (default), Both</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult CreateRelationship(
        string filePath,
        string fromTable,
        string fromColumn,
        string toTable,
        string toColumn,
        bool isActive = true,
        string crossFilterDirection = "Single");

    /// <summary>
    /// Updates an existing relationship's properties
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="fromTable">Source table name</param>
    /// <param name="fromColumn">Source column name</param>
    /// <param name="toTable">Target table name</param>
    /// <param name="toColumn">Target column name</param>
    /// <param name="isActive">New active state (null to keep existing)</param>
    /// <param name="crossFilterDirection">New filter direction (null to keep existing)</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult UpdateRelationship(
        string filePath,
        string fromTable,
        string fromColumn,
        string toTable,
        string toColumn,
        bool? isActive = null,
        string? crossFilterDirection = null);

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
    /// Validates a DAX formula without creating/updating any objects
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="daxFormula">DAX formula to validate</param>
    /// <returns>Result with validation status and any errors</returns>
    DataModelValidationResult ValidateDax(string filePath, string daxFormula);

    /// <summary>
    /// Imports measure definitions from a file
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="importFile">Path to file containing measure definitions (.dax or .json)</param>
    /// <returns>Result indicating success or failure</returns>
    Task<OperationResult> ImportMeasures(string filePath, string importFile);
}
