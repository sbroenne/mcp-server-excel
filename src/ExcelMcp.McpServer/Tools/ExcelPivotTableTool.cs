using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel PivotTable operations
/// </summary>
[McpServerToolType]
public static partial class ExcelPivotTableTool
{
    /// <summary>
    /// PivotTable lifecycle management: create from various sources, list, read details, refresh, and delete.
    ///
    /// BEST PRACTICE: Use 'list' before creating. Prefer 'refresh' or field modifications over delete+recreate.
    /// Delete+recreate loses field configurations, filters, sorting, and custom layouts.
    ///
    /// REFRESH: Call 'refresh' after configuring fields with excel_pivottable_field to update the visual
    /// display. This is especially important for OLAP/Data Model PivotTables where field operations
    /// are structural only and don't automatically trigger a visual refresh.
    ///
    /// LAYOUT (for create operations):
    /// - 0 = Compact (default - row fields in single column with indentation)
    /// - 1 = Tabular (each row field in separate column - best for export/analysis)
    /// - 2 = Outline (hierarchical with expand/collapse)
    ///
    /// CREATE OPTIONS:
    /// - create-from-range: Use sourceSheetName and sourceRangeAddress for data range
    /// - create-from-table: Use sourceTableName for an Excel Table (ListObject)
    /// - create-from-datamodel: Use dataModelTableName for Power Pivot Data Model table
    ///
    /// TIMEOUT: CreateFromDataModel auto-timeouts after 5 minutes for large Data Models.
    ///
    /// RELATED TOOLS:
    /// - excel_pivottable_field: Add/remove/configure fields, filtering, sorting, grouping
    /// - excel_pivottable_calc: Calculated fields, change layout after creation, subtotals
    /// </summary>
    /// <param name="action">The PivotTable operation to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action - required for all operations</param>
    /// <param name="pivotTableName">Name of the PivotTable - required for read, delete, refresh; used as name for create operations</param>
    /// <param name="sourceSheetName">Worksheet name containing source data (for create-from-range)</param>
    /// <param name="sourceRangeAddress">Range address of source data, e.g., 'A1:D100' (for create-from-range)</param>
    /// <param name="sourceTableName">Name of source Excel Table/ListObject (for create-from-table)</param>
    /// <param name="dataModelTableName">Name of Data Model table from Power Pivot (for create-from-datamodel)</param>
    /// <param name="destinationSheetName">Worksheet name where PivotTable will be placed</param>
    /// <param name="destinationCellAddress">Cell address for top-left corner of PivotTable, e.g., 'A3'</param>
    /// <param name="layoutStyle">Layout style for create operations: 0=Compact (default), 1=Tabular, 2=Outline</param>
    [McpServerTool(Name = "excel_pivottable", Title = "Excel PivotTable Operations", Destructive = true)]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelPivotTable(
        PivotTableAction action,
        string sessionId,
        [DefaultValue(null)] string? pivotTableName,
        [DefaultValue(null)] string? sourceSheetName,
        [DefaultValue(null)] string? sourceRangeAddress,
        [DefaultValue(null)] string? sourceTableName,
        [DefaultValue(null)] string? dataModelTableName,
        [DefaultValue(null)] string? destinationSheetName,
        [DefaultValue(null)] string? destinationCellAddress,
        [DefaultValue(null)] int? layoutStyle)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_pivottable",
            ServiceRegistry.PivotTable.ToActionString(action),
            () => action switch
            {
                PivotTableAction.List => ForwardList(sessionId),
                PivotTableAction.Read => ForwardRead(sessionId, pivotTableName),
                PivotTableAction.CreateFromRange => ForwardCreateFromRange(sessionId, sourceSheetName, sourceRangeAddress, destinationSheetName, destinationCellAddress, pivotTableName),
                PivotTableAction.CreateFromTable => ForwardCreateFromTable(sessionId, sourceTableName, destinationSheetName, destinationCellAddress, pivotTableName),
                PivotTableAction.CreateFromDataModel => ForwardCreateFromDataModel(sessionId, dataModelTableName, destinationSheetName, destinationCellAddress, pivotTableName),
                PivotTableAction.Delete => ForwardDelete(sessionId, pivotTableName),
                PivotTableAction.Refresh => ForwardRefresh(sessionId, pivotTableName),
                _ => throw new ArgumentException($"Unknown action: {action} ({ServiceRegistry.PivotTable.ToActionString(action)})", nameof(action))
            });
    }

    // === SERVICE FORWARDING METHODS ===

    private static string ForwardList(string sessionId)
    {
        return ExcelToolsBase.ForwardToService("pivottable.list", sessionId);
    }

    private static string ForwardRead(string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "read");

        return ExcelToolsBase.ForwardToService("pivottable.read", sessionId, new { pivotTableName });
    }

    private static string ForwardCreateFromRange(
        string sessionId,
        string? sourceSheetName,
        string? sourceRangeAddress,
        string? destinationSheetName,
        string? destinationCellAddress,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(sourceSheetName))
            ExcelToolsBase.ThrowMissingParameter("sourceSheetName", "create-from-range");
        if (string.IsNullOrWhiteSpace(sourceRangeAddress))
            ExcelToolsBase.ThrowMissingParameter("sourceRangeAddress", "create-from-range");
        if (string.IsNullOrWhiteSpace(destinationSheetName))
            ExcelToolsBase.ThrowMissingParameter("destinationSheetName", "create-from-range");
        if (string.IsNullOrWhiteSpace(destinationCellAddress))
            ExcelToolsBase.ThrowMissingParameter("destinationCellAddress", "create-from-range");
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "create-from-range");

        return ExcelToolsBase.ForwardToService("pivottable.create-from-range", sessionId, new
        {
            sourceSheet = sourceSheetName,
            sourceRange = sourceRangeAddress,
            destinationSheet = destinationSheetName,
            destinationCell = destinationCellAddress,
            pivotTableName
        });
    }

    private static string ForwardCreateFromTable(
        string sessionId,
        string? sourceTableName,
        string? destinationSheetName,
        string? destinationCellAddress,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(sourceTableName))
            ExcelToolsBase.ThrowMissingParameter("sourceTableName", "create-from-table");
        if (string.IsNullOrWhiteSpace(destinationSheetName))
            ExcelToolsBase.ThrowMissingParameter("destinationSheetName", "create-from-table");
        if (string.IsNullOrWhiteSpace(destinationCellAddress))
            ExcelToolsBase.ThrowMissingParameter("destinationCellAddress", "create-from-table");
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "create-from-table");

        return ExcelToolsBase.ForwardToService("pivottable.create-from-table", sessionId, new
        {
            tableName = sourceTableName,
            destinationSheet = destinationSheetName,
            destinationCell = destinationCellAddress,
            pivotTableName
        });
    }

    private static string ForwardCreateFromDataModel(
        string sessionId,
        string? dataModelTableName,
        string? destinationSheetName,
        string? destinationCellAddress,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(dataModelTableName))
            ExcelToolsBase.ThrowMissingParameter("dataModelTableName", "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(destinationSheetName))
            ExcelToolsBase.ThrowMissingParameter("destinationSheetName", "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(destinationCellAddress))
            ExcelToolsBase.ThrowMissingParameter("destinationCellAddress", "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "create-from-datamodel");

        return ExcelToolsBase.ForwardToService("pivottable.create-from-datamodel", sessionId, new
        {
            tableName = dataModelTableName,
            destinationSheet = destinationSheetName,
            destinationCell = destinationCellAddress,
            pivotTableName
        });
    }

    private static string ForwardDelete(string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "delete");

        return ExcelToolsBase.ForwardToService("pivottable.delete", sessionId, new { pivotTableName });
    }

    private static string ForwardRefresh(string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "refresh");

        return ExcelToolsBase.ForwardToService("pivottable.refresh", sessionId, new { pivotTableName });
    }
}




