using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Data Model (Power Pivot) - tables and DAX measures.
/// Use excel_datamodel_rel for relationships.
/// </summary>
[McpServerToolType]
public static partial class ExcelDataModelTool
{
    /// <summary>
    /// Data Model (Power Pivot) - DAX measures and table management.
    ///
    /// CRITICAL: WORKSHEET TABLES AND DATA MODEL ARE SEPARATE!
    /// - After excel_table(append) changes, Data Model still has OLD data
    /// - MUST call excel_datamodel(refresh) to sync changes
    /// - Power Query refresh auto-syncs (no manual refresh needed)
    ///
    /// PREREQUISITE: Tables must be added to the Data Model first.
    /// Use excel_table with add-to-datamodel action to add worksheet tables,
    /// or excel_powerquery to import and load data directly to the Data Model.
    ///
    /// DAX MEASURES:
    /// - Create measures with DAX formulas like 'SUM(Sales[Amount])'
    /// - Measures can reference columns, other measures, and use DAX functions
    /// - Format string uses US format codes like '#,##0.00' for currency
    /// - DAX formulas are auto-formatted on CREATE/UPDATE via Dax.Formatter library (SQLBI)
    /// - Read operations (list-measures, read) return raw DAX as stored
    ///
    /// DAX EVALUATE QUERIES:
    /// - Use 'evaluate' action to execute DAX EVALUATE queries against the Data Model
    /// - Returns tabular results (columns and rows) from queries like 'EVALUATE TableName'
    /// - Supports complex DAX: SUMMARIZE, FILTER, CALCULATETABLE, TOPN, etc.
    /// - Results can be used for analysis or to create DAX-backed tables via excel_table
    ///
    /// DMV (DYNAMIC MANAGEMENT VIEW) QUERIES:
    /// - Use 'execute-dmv' action to query Data Model metadata via SQL-like syntax
    /// - Syntax: SELECT * FROM $SYSTEM.SchemaRowset (ONLY SELECT * supported, not specific columns)
    /// - WORKING DMVs: TMSCHEMA_MEASURES, TMSCHEMA_RELATIONSHIPS, DISCOVER_CALC_DEPENDENCY, DBSCHEMA_CATALOGS
    /// - LIMITED DMVs: TMSCHEMA_TABLES, TMSCHEMA_COLUMNS may return empty in Excel's embedded AS
    /// - Use DISCOVER_SCHEMA_ROWSETS to list all available DMVs
    /// - See MS docs: https://learn.microsoft.com/analysis-services/instances/use-dynamic-management-views-dmvs
    ///
    /// DAX FILE INPUT:
    /// - daxFormula, daxQuery, dmvQuery accept inline formulas OR file paths
    /// - Use daxFormulaFile, daxQueryFile, dmvQueryFile to load from files
    /// - Useful for complex multi-line DAX that's hard to pass on command line
    ///
    /// PERFORMANCE:
    /// - Formatting adds ~100-500ms network latency per operation (daxformatter.com API)
    /// - Graceful fallback: returns original DAX if formatting fails (no operation failures)
    ///
    /// DESTRUCTIVE OPERATIONS:
    /// - delete-table: Removes table AND all its measures - cannot be undone
    ///
    /// TIMEOUT: Operations auto-timeout after 2 minutes for large Data Models.
    ///
    /// RELATED TOOLS:
    /// - excel_table: Add worksheet tables to Data Model (add-to-datamodel action)
    /// - excel_table: Create DAX-backed tables (create-from-dax action)
    /// - excel_datamodel_rel: Manage relationships between tables
    /// </summary>
    /// <param name="action">The Data Model operation to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action</param>
    /// <param name="measureName">Name of the DAX measure - required for measure operations</param>
    /// <param name="tableName">Name of the table in the Data Model - required for table operations and create-measure</param>
    /// <param name="newName">New name for rename-table action</param>
    /// <param name="daxFormula">DAX formula for the measure, e.g., 'SUM(Sales[Amount])'</param>
    /// <param name="daxFormulaFile">Path to file containing DAX formula (alternative to daxFormula)</param>
    /// <param name="daxQuery">DAX EVALUATE query for evaluate action, e.g., 'EVALUATE TableName'</param>
    /// <param name="daxQueryFile">Path to file containing DAX query (alternative to daxQuery)</param>
    /// <param name="dmvQuery">DMV query for execute-dmv action. Use SELECT * FROM $SYSTEM.TMSCHEMA_MEASURES or DISCOVER_CALC_DEPENDENCY</param>
    /// <param name="dmvQueryFile">Path to file containing DMV query (alternative to dmvQuery)</param>
    /// <param name="description">Optional description for the measure</param>
    /// <param name="formatType">Number format code in US format, e.g., '#,##0.00' for currency</param>
    /// <param name="timeout">Optional timeout for refresh operation</param>
    [McpServerTool(Name = "excel_datamodel", Title = "Excel Data Model Operations", Destructive = true)]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelDataModel(
        DataModelAction action,
        string sessionId,
        [DefaultValue(null)] string? measureName,
        [DefaultValue(null)] string? tableName,
        [DefaultValue(null)] string? newName,
        [DefaultValue(null)] string? daxFormula,
        [DefaultValue(null)] string? daxFormulaFile,
        [DefaultValue(null)] string? daxQuery,
        [DefaultValue(null)] string? daxQueryFile,
        [DefaultValue(null)] string? dmvQuery,
        [DefaultValue(null)] string? dmvQueryFile,
        [DefaultValue(null)] string? description,
        [DefaultValue(null)] string? formatType,
        [DefaultValue(null)] TimeSpan? timeout)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_datamodel",
            ServiceRegistry.DataModel.ToActionString(action),
            () => ServiceRegistry.DataModel.RouteAction(
                action,
                sessionId,
                ExcelToolsBase.ForwardToServiceFunc,
                tableName: tableName,
                measureName: measureName,
                oldName: tableName,  // rename-table uses oldName for source
                newName: newName,
                timeout: timeout,
                daxFormula: daxFormula,
                daxFormulaFile: daxFormulaFile,
                formatType: formatType,
                description: description,
                daxQuery: daxQuery,
                daxQueryFile: daxQueryFile,
                dmvQuery: dmvQuery,
                dmvQueryFile: dmvQueryFile));
    }
}




