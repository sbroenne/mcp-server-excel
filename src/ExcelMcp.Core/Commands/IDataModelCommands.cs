using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;

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
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <returns>Result containing list of tables with metadata</returns>
    Task<DataModelTableListResult> ListTablesAsync(IExcelBatch batch);

    /// <summary>
    /// Lists all DAX measures in the model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Optional: Filter measures by table name</param>
    /// <returns>Result containing list of measures with formulas</returns>
    Task<DataModelMeasureListResult> ListMeasuresAsync(IExcelBatch batch, string? tableName = null);

    /// <summary>
    /// Views complete measure details and DAX formula
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="measureName">Name of the measure to view</param>
    /// <returns>Result containing complete measure information</returns>
    Task<DataModelMeasureViewResult> ViewMeasureAsync(IExcelBatch batch, string measureName);

    /// <summary>
    /// Exports measure DAX formula to file with metadata
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="measureName">Name of the measure to export</param>
    /// <param name="outputFile">Path to output DAX file</param>
    /// <returns>Result indicating success or failure</returns>
    Task<OperationResult> ExportMeasureAsync(IExcelBatch batch, string measureName, string outputFile);

    /// <summary>
    /// Lists all table relationships in the model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <returns>Result containing list of relationships</returns>
    Task<DataModelRelationshipListResult> ListRelationshipsAsync(IExcelBatch batch);

    /// <summary>
    /// Deletes a DAX measure from the Data Model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="measureName">Name of the measure to delete</param>
    /// <returns>Result indicating success or failure</returns>
    Task<OperationResult> DeleteMeasureAsync(IExcelBatch batch, string measureName);

    /// <summary>
    /// Deletes a relationship from the Data Model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="fromTable">Source table name</param>
    /// <param name="fromColumn">Source column name</param>
    /// <param name="toTable">Target table name</param>
    /// <param name="toColumn">Target column name</param>
    /// <returns>Result indicating success or failure</returns>
    Task<OperationResult> DeleteRelationshipAsync(IExcelBatch batch, string fromTable, string fromColumn, string toTable, string toColumn);

    /// <summary>
    /// Refreshes entire Data Model or specific table
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Optional: Specific table to refresh (if null, refreshes entire model)</param>
    /// <returns>Result indicating success or failure</returns>
    Task<OperationResult> RefreshAsync(IExcelBatch batch, string? tableName = null);
}
