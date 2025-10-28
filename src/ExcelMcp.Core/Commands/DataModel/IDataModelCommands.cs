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
    /// Lists all columns in a Data Model table
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Name of the table to list columns from</param>
    /// <returns>Result containing list of columns with metadata</returns>
    Task<DataModelTableColumnsResult> ListTableColumnsAsync(IExcelBatch batch, string tableName);

    /// <summary>
    /// Views complete table details including columns and measures
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Name of the table to view</param>
    /// <returns>Result containing complete table information</returns>
    Task<DataModelTableViewResult> ViewTableAsync(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets overall Data Model summary statistics
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <returns>Result containing model metadata (table count, measure count, etc.)</returns>
    Task<DataModelInfoResult> GetModelInfoAsync(IExcelBatch batch);

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

    /// <summary>
    /// Creates a new DAX measure in the Data Model
    /// Uses Excel COM API: ModelMeasures.Add method (Office 2016+)
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Name of the table to add the measure to</param>
    /// <param name="measureName">Name of the new measure</param>
    /// <param name="daxFormula">DAX formula for the measure</param>
    /// <param name="formatType">Optional: Format type (Currency, Decimal, Percentage, General)</param>
    /// <param name="description">Optional: Description of the measure</param>
    /// <returns>Result indicating success or failure</returns>
    Task<OperationResult> CreateMeasureAsync(IExcelBatch batch, string tableName, string measureName,
                                             string daxFormula, string? formatType = null,
                                             string? description = null);

    /// <summary>
    /// Updates an existing DAX measure in the Data Model
    /// Uses Excel COM API: ModelMeasure properties (Formula, Description, FormatInformation - all Read/Write)
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="measureName">Name of the measure to update</param>
    /// <param name="daxFormula">Optional: New DAX formula (null to keep existing)</param>
    /// <param name="formatType">Optional: New format type (null to keep existing)</param>
    /// <param name="description">Optional: New description (null to keep existing)</param>
    /// <returns>Result indicating success or failure</returns>
    Task<OperationResult> UpdateMeasureAsync(IExcelBatch batch, string measureName,
                                             string? daxFormula = null, string? formatType = null,
                                             string? description = null);

    /// <summary>
    /// Creates a new relationship between two tables in the Data Model
    /// Uses Excel COM API: ModelRelationships.Add method (Office 2016+)
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="fromTable">Source table name</param>
    /// <param name="fromColumn">Source column name</param>
    /// <param name="toTable">Target table name</param>
    /// <param name="toColumn">Target column name</param>
    /// <param name="active">Whether the relationship should be active (default: true)</param>
    /// <returns>Result indicating success or failure</returns>
    Task<OperationResult> CreateRelationshipAsync(IExcelBatch batch, string fromTable,
                                                   string fromColumn, string toTable,
                                                   string toColumn, bool active = true);

    /// <summary>
    /// Updates an existing relationship's active state in the Data Model
    /// Uses Excel COM API: ModelRelationship.Active property (Read/Write)
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="fromTable">Source table name</param>
    /// <param name="fromColumn">Source column name</param>
    /// <param name="toTable">Target table name</param>
    /// <param name="toColumn">Target column name</param>
    /// <param name="active">New active state for the relationship</param>
    /// <returns>Result indicating success or failure</returns>
    Task<OperationResult> UpdateRelationshipAsync(IExcelBatch batch, string fromTable,
                                                   string fromColumn, string toTable,
                                                   string toColumn, bool active);
}
