using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Data Model relationship operations.
/// Use excel_datamodel for tables and DAX measures.
/// </summary>
[McpServerToolType]
public static partial class ExcelDataModelRelTool
{
    /// <summary>
    /// Data Model relationships - link tables for cross-table DAX calculations.
    ///
    /// CRITICAL: Deleting or recreating tables (via excel_datamodel delete-table or excel_table delete)
    /// removes ALL their relationships. Use 'list' before table operations to backup,
    /// then recreate relationships after schema changes.
    ///
    /// BEST PRACTICE: Use 'list' to check existing relationships before creating.
    ///
    /// RELATIONSHIP REQUIREMENTS:
    /// - Both tables must exist in the Data Model first
    /// - Columns must have compatible data types
    /// - From column is typically the many-side (detail table)
    /// - To column is typically the one-side (lookup table)
    ///
    /// ACTIVE VS INACTIVE:
    /// - Only ONE active relationship can exist between two tables
    /// - Use active=false when creating alternative paths
    /// - DAX USERELATIONSHIP() activates inactive relationships
    ///
    /// TIMEOUT: Operations auto-timeout after 2 minutes for large Data Models.
    ///
    /// RELATED TOOLS:
    /// - excel_datamodel: Tables, measures, Data Model info
    /// - excel_table: Add tables to Data Model (add-to-datamodel action)
    /// </summary>
    /// <param name="action">The relationship operation to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action</param>
    /// <param name="fromTable">Source (many-side) table name containing the foreign key</param>
    /// <param name="fromColumn">Column in fromTable that links to toTable</param>
    /// <param name="toTable">Target (one-side/lookup) table name containing the primary key</param>
    /// <param name="toColumn">Column in toTable that fromColumn links to (usually primary key)</param>
    /// <param name="active">True for active relationship (default), false for inactive (use with DAX USERELATIONSHIP)</param>
    [McpServerTool(Name = "excel_datamodel_rel", Title = "Excel Data Model Relationship Operations", Destructive = true)]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelDataModelRel(
        DataModelRelAction action,
        string sessionId,
        [DefaultValue(null)] string? fromTable,
        [DefaultValue(null)] string? fromColumn,
        [DefaultValue(null)] string? toTable,
        [DefaultValue(null)] string? toColumn,
        [DefaultValue(true)] bool? active)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_datamodel_rel",
            ServiceRegistry.DataModelRel.ToActionString(action),
            () => ServiceRegistry.DataModelRel.RouteAction(
                action,
                sessionId,
                ExcelToolsBase.ForwardToServiceFunc,
                fromTable: fromTable,
                fromColumn: fromColumn,
                toTable: toTable,
                toColumn: toColumn,
                active: active));
    }
}




