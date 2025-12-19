using Sbroenne.ExcelMcp.ComInterop.Session;
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
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <returns>Result containing list of tables with metadata</returns>
    DataModelTableListResult ListTables(IExcelBatch batch);

    /// <summary>
    /// Lists all columns in a Data Model table
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Name of the table to list columns from</param>
    /// <returns>Result containing list of columns with metadata</returns>
    DataModelTableColumnsResult ListColumns(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets complete table details including columns and measures
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Name of the table to get</param>
    /// <returns>Result containing complete table information</returns>
    DataModelTableViewResult ReadTable(IExcelBatch batch, string tableName);

    /// <summary>
    /// Gets overall Data Model summary statistics
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <returns>Result containing model metadata (table count, measure count, etc.)</returns>
    DataModelInfoResult ReadInfo(IExcelBatch batch);

    /// <summary>
    /// Lists all DAX measures in the model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Optional: Filter measures by table name</param>
    /// <returns>Result containing list of measures with formulas</returns>
    DataModelMeasureListResult ListMeasures(IExcelBatch batch, string? tableName = null);

    /// <summary>
    /// Gets complete measure details and DAX formula
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="measureName">Name of the measure to get</param>
    /// <returns>Result containing complete measure information</returns>
    DataModelMeasureViewResult Read(IExcelBatch batch, string measureName);

    /// <summary>
    /// Lists all table relationships in the model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <returns>Result containing list of relationships</returns>
    DataModelRelationshipListResult ListRelationships(IExcelBatch batch);

    /// <summary>
    /// Gets a specific relationship by its table/column identifiers
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="fromTable">Source table name</param>
    /// <param name="fromColumn">Source column name</param>
    /// <param name="toTable">Target table name</param>
    /// <param name="toColumn">Target column name</param>
    /// <returns>Result containing relationship details</returns>
    DataModelRelationshipViewResult ReadRelationship(IExcelBatch batch, string fromTable, string fromColumn, string toTable, string toColumn);

    /// <summary>
    /// Deletes a DAX measure from the Data Model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="measureName">Name of the measure to delete</param>
    /// <exception cref="ArgumentException">Thrown when measureName is invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when measure not found or deletion fails</exception>
    void DeleteMeasure(IExcelBatch batch, string measureName);

    /// <summary>
    /// Deletes a table from the Data Model.
    /// Use this to remove orphaned tables created when Power Query is deleted and recreated
    /// with a different name, leaving stale tables in the Data Model.
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Name of the table to delete</param>
    /// <exception cref="ArgumentException">Thrown when tableName is invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when table not found or deletion fails</exception>
    void DeleteTable(IExcelBatch batch, string tableName);

    /// <summary>
    /// Renames a table in the Data Model.
    /// Names are trimmed before comparison; a no-op success is returned when
    /// trimmed old and new names match (including case-only change no-op).
    /// Case-only renames are allowed if new name differs only in casing.
    /// Conflict detection is case-insensitive, excluding the target table.
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="oldName">Current name of the table</param>
    /// <param name="newName">New name for the table</param>
    /// <returns>RenameResult with ObjectType="data-model-table"</returns>
    RenameResult RenameTable(IExcelBatch batch, string oldName, string newName);

    /// <summary>
    /// Deletes a relationship from the Data Model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="fromTable">Source table name</param>
    /// <param name="fromColumn">Source column name</param>
    /// <param name="toTable">Target table name</param>
    /// <param name="toColumn">Target column name</param>
    /// <exception cref="ArgumentException">Thrown when parameters are invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when relationship not found or deletion fails</exception>
    void DeleteRelationship(IExcelBatch batch, string fromTable, string fromColumn, string toTable, string toColumn);

    /// <summary>
    /// Refreshes entire Data Model or specific table
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Optional: Specific table to refresh (if null, refreshes entire model)</param>
    /// <exception cref="InvalidOperationException">Thrown when refresh operation fails</exception>
    void Refresh(IExcelBatch batch, string? tableName = null);

    /// <summary>
    /// Refreshes Data Model table(s) with timeout
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Optional: Specific table to refresh (if null, refreshes entire model)</param>
    /// <param name="timeout">Timeout for the refresh operation</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when timeout is invalid</exception>
    /// <exception cref="OperationCanceledException">Thrown when operation times out</exception>
    /// <exception cref="InvalidOperationException">Thrown when refresh operation fails</exception>
    void Refresh(IExcelBatch batch, string? tableName, TimeSpan? timeout);

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
    /// <exception cref="ArgumentException">Thrown when parameters are invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when table not found or creation fails</exception>
    void CreateMeasure(IExcelBatch batch, string tableName, string measureName,
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
    /// <exception cref="ArgumentException">Thrown when measureName is invalid or all parameters are null</exception>
    /// <exception cref="InvalidOperationException">Thrown when measure not found or update fails</exception>
    void UpdateMeasure(IExcelBatch batch, string measureName,
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
    /// <exception cref="ArgumentException">Thrown when parameters are invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when tables/columns not found or creation fails</exception>
    void CreateRelationship(IExcelBatch batch, string fromTable,
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
    /// <exception cref="ArgumentException">Thrown when parameters are invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when relationship not found or update fails</exception>
    void UpdateRelationship(IExcelBatch batch, string fromTable,
                            string fromColumn, string toTable,
                            string toColumn, bool active);
}

