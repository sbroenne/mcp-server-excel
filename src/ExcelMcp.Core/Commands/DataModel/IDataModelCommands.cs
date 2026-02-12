using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model (Power Pivot) - DAX measures and table management.
///
/// CRITICAL: WORKSHEET TABLES AND DATA MODEL ARE SEPARATE!
/// - After table append changes, Data Model still has OLD data
/// - MUST call refresh to sync changes
/// - Power Query refresh auto-syncs (no manual refresh needed)
///
/// PREREQUISITE: Tables must be added to the Data Model first.
/// Use table add-to-datamodel for worksheet tables,
/// or powerquery to import and load data directly to the Data Model.
///
/// DAX MEASURES:
/// - Create with DAX formulas like 'SUM(Sales[Amount])'
/// - DAX formulas are auto-formatted on CREATE/UPDATE via Dax.Formatter (SQLBI)
/// - Read operations return raw DAX as stored
///
/// DAX EVALUATE QUERIES:
/// - Use evaluate to execute DAX EVALUATE queries against the Data Model
/// - Returns tabular results from queries like 'EVALUATE TableName'
/// - Supports complex DAX: SUMMARIZE, FILTER, CALCULATETABLE, TOPN, etc.
///
/// DMV (DYNAMIC MANAGEMENT VIEW) QUERIES:
/// - Use execute-dmv to query Data Model metadata via SQL-like syntax
/// - Syntax: SELECT * FROM $SYSTEM.SchemaRowset (ONLY SELECT * supported)
/// - Use DISCOVER_SCHEMA_ROWSETS to list all available DMVs
///
/// Use datamodelrel for relationships between tables.
/// </summary>
[ServiceCategory("datamodel", "DataModel")]
[McpTool("datamodel", Title = "Data Model Operations", Destructive = true, Category = "analysis",
    Description = "Data Model (Power Pivot) - DAX measures and table management. CRITICAL: Worksheet tables and Data Model are separate! After table(append), MUST call datamodel(refresh) to sync. Power Query refresh auto-syncs. DAX MEASURES: Create with formulas like SUM(Sales[Amount]), auto-formatted via daxformatter.com. DAX EVALUATE: Execute queries (SUMMARIZE, FILTER, CALCULATETABLE, TOPN). DMV QUERIES: SELECT * FROM $SYSTEM.SchemaRowset for metadata. DAX FILE INPUT: daxFormulaFile/daxQueryFile for complex multi-line DAX. TIMEOUT: 2 min. Use datamodel_relationship for relationships, table for add-to-datamodel.")]
public interface IDataModelCommands
{
    /// <summary>
    /// Lists all tables in the Data Model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <returns>Result containing list of tables with metadata</returns>
    [ServiceAction("list-tables")]
    DataModelTableListResult ListTables(IExcelBatch batch);

    /// <summary>
    /// Lists all columns in a Data Model table
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Name of the table to list columns from</param>
    /// <returns>Result containing list of columns with metadata</returns>
    [ServiceAction("list-columns")]
    DataModelTableColumnsResult ListColumns(IExcelBatch batch, [RequiredParameter] string tableName);

    /// <summary>
    /// Gets complete table details including columns and measures
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Name of the table to get</param>
    /// <returns>Result containing complete table information</returns>
    [ServiceAction("read-table")]
    DataModelTableViewResult ReadTable(IExcelBatch batch, [RequiredParameter] string tableName);

    /// <summary>
    /// Gets overall Data Model summary statistics
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <returns>Result containing model metadata (table count, measure count, etc.)</returns>
    [ServiceAction("read-info")]
    DataModelInfoResult ReadInfo(IExcelBatch batch);

    /// <summary>
    /// Lists all DAX measures in the model.
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Optional: Filter measures by table name</param>
    /// <returns>Result containing list of measures with formula previews</returns>
    [ServiceAction("list-measures")]
    DataModelMeasureListResult ListMeasures(IExcelBatch batch, string? tableName = null);

    /// <summary>
    /// Gets complete measure details and DAX formula.
    /// Returns the raw DAX formula as stored in the Data Model.
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="measureName">Name of the measure to get</param>
    /// <returns>Result containing complete measure information with DAX formula</returns>
    [ServiceAction("read")]
    DataModelMeasureViewResult Read(IExcelBatch batch, [RequiredParameter] string measureName);

    /// <summary>
    /// Deletes a DAX measure from the Data Model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="measureName">Name of the measure to delete</param>
    /// <exception cref="ArgumentException">Thrown when measureName is invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when measure not found or deletion fails</exception>
    [ServiceAction("delete-measure")]
    void DeleteMeasure(IExcelBatch batch, [RequiredParameter] string measureName);

    /// <summary>
    /// Deletes a table from the Data Model.
    /// Use this to remove orphaned tables created when Power Query is deleted and recreated
    /// with a different name, leaving stale tables in the Data Model.
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Name of the table to delete</param>
    /// <exception cref="ArgumentException">Thrown when tableName is invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when table not found or deletion fails</exception>
    [ServiceAction("delete-table")]
    void DeleteTable(IExcelBatch batch, [RequiredParameter] string tableName);

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
    [ServiceAction("rename-table")]
    RenameResult RenameTable(IExcelBatch batch, [RequiredParameter] string oldName, [RequiredParameter] string newName);

    /// <summary>
    /// Refreshes entire Data Model or specific table
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Optional: Specific table to refresh (if null, refreshes entire model)</param>
    /// <param name="timeout">Optional: Timeout for the refresh operation</param>
    /// <exception cref="InvalidOperationException">Thrown when refresh operation fails</exception>
    [ServiceAction("refresh")]
    void Refresh(IExcelBatch batch, string? tableName = null, TimeSpan? timeout = null);

    /// <summary>
    /// Creates a new DAX measure in the Data Model.
    /// DAX formula is automatically formatted with proper indentation before saving.
    /// Uses Excel COM API: ModelMeasures.Add method (Office 2016+)
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="tableName">Name of the table to add the measure to</param>
    /// <param name="measureName">Name of the new measure</param>
    /// <param name="daxFormula">DAX formula for the measure (will be auto-formatted)</param>
    /// <param name="formatType">Optional: Format type (Currency, Decimal, Percentage, General)</param>
    /// <param name="description">Optional: Description of the measure</param>
    /// <exception cref="ArgumentException">Thrown when parameters are invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when table not found or creation fails</exception>
    [ServiceAction("create-measure")]
    void CreateMeasure(
        IExcelBatch batch,
        [RequiredParameter] string tableName,
        [RequiredParameter] string measureName,
        [RequiredParameter, FileOrValue] string daxFormula,
        string? formatType = null,
        string? description = null);

    /// <summary>
    /// Updates an existing DAX measure in the Data Model.
    /// DAX formula is automatically formatted with proper indentation before saving.
    /// Uses Excel COM API: ModelMeasure properties (Formula, Description, FormatInformation - all Read/Write)
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="measureName">Name of the measure to update</param>
    /// <param name="daxFormula">Optional: New DAX formula (null to keep existing, will be auto-formatted if provided)</param>
    /// <param name="formatType">Optional: New format type (null to keep existing)</param>
    /// <param name="description">Optional: New description (null to keep existing)</param>
    /// <exception cref="ArgumentException">Thrown when measureName is invalid or all parameters are null</exception>
    /// <exception cref="InvalidOperationException">Thrown when measure not found or update fails</exception>
    [ServiceAction("update-measure")]
    void UpdateMeasure(
        IExcelBatch batch,
        [RequiredParameter] string measureName,
        [FileOrValue] string? daxFormula = null,
        string? formatType = null,
        string? description = null);

    /// <summary>
    /// Executes a DAX EVALUATE query against the Data Model and returns the results.
    /// Uses ADOConnection.Execute for direct DAX query execution via MSOLAP provider.
    /// The query should start with EVALUATE and return a table result.
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="daxQuery">DAX EVALUATE query (e.g., "EVALUATE 'TableName'" or "EVALUATE SUMMARIZE(...)")</param>
    /// <returns>Result containing column names and data rows from the DAX query</returns>
    /// <exception cref="ArgumentException">Thrown when daxQuery is empty</exception>
    /// <exception cref="InvalidOperationException">Thrown when workbook has no Data Model or query execution fails</exception>
    [ServiceAction("evaluate")]
    DaxEvaluateResult Evaluate(IExcelBatch batch, [RequiredParameter, FileOrValue] string daxQuery);

    /// <summary>
    /// Executes a DMV (Dynamic Management View) query against the Data Model and returns the results.
    /// Uses ADOConnection.Execute for SQL-like DMV query execution via MSOLAP provider.
    /// DMV queries retrieve metadata about the Data Model (tables, columns, measures, relationships, etc.).
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="dmvQuery">DMV query in SQL-like syntax (e.g., "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES")</param>
    /// <returns>Result containing column names and data rows from the DMV query</returns>
    /// <exception cref="ArgumentException">Thrown when dmvQuery is empty</exception>
    /// <exception cref="InvalidOperationException">Thrown when workbook has no Data Model or query execution fails</exception>
    /// <remarks>
    /// Common DMV queries for Excel PowerPivot:
    /// - $SYSTEM.TMSCHEMA_TABLES - List all tables
    /// - $SYSTEM.TMSCHEMA_COLUMNS - List all columns
    /// - $SYSTEM.TMSCHEMA_MEASURES - List all measures
    /// - $SYSTEM.TMSCHEMA_RELATIONSHIPS - List all relationships
    /// - $SYSTEM.DISCOVER_CALC_DEPENDENCY - Show calculation dependencies
    /// </remarks>
    [ServiceAction("execute-dmv")]
    DmvQueryResult ExecuteDmv(IExcelBatch batch, [RequiredParameter, FileOrValue] string dmvQuery);
}



