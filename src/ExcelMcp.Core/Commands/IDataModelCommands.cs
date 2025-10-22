using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model management commands - Basic operations using Excel COM API
/// Provides read-only access to Data Model tables, measures, and relationships
/// </summary>
public interface IDataModelCommands
{
    /// <summary>
    /// Lists all tables in the Data Model
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <returns>Result containing list of tables with metadata</returns>
    DataModelTableListResult ListTables(string filePath);

    /// <summary>
    /// Lists all DAX measures in the model
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="tableName">Optional: Filter measures by table name</param>
    /// <returns>Result containing list of measures with formulas</returns>
    DataModelMeasureListResult ListMeasures(string filePath, string? tableName = null);

    /// <summary>
    /// Views complete measure details and DAX formula
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="measureName">Name of the measure to view</param>
    /// <returns>Result containing complete measure information</returns>
    DataModelMeasureViewResult ViewMeasure(string filePath, string measureName);

    /// <summary>
    /// Exports measure DAX formula to file with metadata
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="measureName">Name of the measure to export</param>
    /// <param name="outputFile">Path to output DAX file</param>
    /// <returns>Result indicating success or failure</returns>
    Task<OperationResult> ExportMeasure(string filePath, string measureName, string outputFile);

    /// <summary>
    /// Lists all table relationships in the model
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <returns>Result containing list of relationships</returns>
    DataModelRelationshipListResult ListRelationships(string filePath);

    /// <summary>
    /// Deletes a DAX measure from the Data Model
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="measureName">Name of the measure to delete</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult DeleteMeasure(string filePath, string measureName);

    /// <summary>
    /// Deletes a relationship from the Data Model
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="fromTable">Source table name</param>
    /// <param name="fromColumn">Source column name</param>
    /// <param name="toTable">Target table name</param>
    /// <param name="toColumn">Target column name</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult DeleteRelationship(string filePath, string fromTable, string fromColumn, string toTable, string toColumn);

    /// <summary>
    /// Refreshes entire Data Model or specific table
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="tableName">Optional: Specific table to refresh (if null, refreshes entire model)</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult Refresh(string filePath, string? tableName = null);
}
