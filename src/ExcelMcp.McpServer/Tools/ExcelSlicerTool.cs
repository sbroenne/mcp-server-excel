using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Table;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Slicer operations
/// </summary>
[McpServerToolType]
public static class ExcelSlicerTool
{
    private static readonly JsonSerializerOptions JsonOptions = ExcelToolsBase.JsonOptions;

    /// <summary>
    /// Slicer management: create, list, configure, and delete visual filtering controls.
    ///
    /// SLICERS provide visual filtering for PivotTables AND Tables:
    /// - Click items in the slicer to filter connected data
    /// - PivotTable slicers can filter multiple PivotTables
    /// - Table slicers filter a single Excel Table
    ///
    /// PIVOTTABLE SLICERS:
    /// - create-slicer: Requires pivotTableName, fieldName, slicerName, destinationSheet, position
    /// - list-slicers: Returns PivotTable slicers. Optionally filter by pivotTableName
    /// - set-slicer-selection: Pass slicerName and selectedItems (JSON array)
    /// - delete-slicer: Requires slicerName
    ///
    /// TABLE SLICERS:
    /// - create-table-slicer: Requires tableName, columnName, slicerName, destinationSheet, position
    /// - list-table-slicers: Returns Table slicers. Optionally filter by tableName
    /// - set-table-slicer-selection: Pass slicerName and selectedItems (JSON array)
    /// - delete-table-slicer: Requires slicerName
    ///
    /// SELECTION: Pass selectedItems as JSON array of strings.
    /// Empty array clears filter (shows all items). Set clearFirst=false to add
    /// to existing selection instead of replacing.
    ///
    /// RELATED TOOLS:
    /// - excel_pivottable: Create PivotTables (required before creating PivotTable slicers)
    /// - excel_pivottable_field: Manage PivotTable fields (use list-fields to see available field names)
    /// - excel_table: Create and manage Excel Tables (required before creating Table slicers)
    /// </summary>
    /// <param name="action">The slicer operation to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action</param>
    /// <param name="pivotTableName">Name of the PivotTable - required for create-slicer, optional for list-slicers (filters results)</param>
    /// <param name="tableName">Name of the Excel Table - required for create-table-slicer, optional for list-table-slicers (filters results)</param>
    /// <param name="fieldName">Field name from the PivotTable to create slicer for (required for create-slicer)</param>
    /// <param name="columnName">Column name from the Table to create slicer for (required for create-table-slicer)</param>
    /// <param name="slicerName">Name for the slicer - required for create/delete/set-selection actions</param>
    /// <param name="destinationSheet">Worksheet where slicer will be placed (required for create actions)</param>
    /// <param name="position">Cell address for top-left corner of slicer, e.g., 'E1' (required for create actions)</param>
    /// <param name="selectedItems">JSON array of item names to select, e.g., '["North","South"]'. Empty array clears filter.</param>
    /// <param name="clearFirst">If true (default), replaces selection. If false, adds to existing selection.</param>
    [McpServerTool(Name = "excel_slicer", Title = "Excel Slicer Operations")]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static string ExcelSlicer(
        SlicerAction action,
        string sessionId,
        [DefaultValue(null)] string? pivotTableName,
        [DefaultValue(null)] string? tableName,
        [DefaultValue(null)] string? fieldName,
        [DefaultValue(null)] string? columnName,
        [DefaultValue(null)] string? slicerName,
        [DefaultValue(null)] string? destinationSheet,
        [DefaultValue(null)] string? position,
        [DefaultValue(null)] string? selectedItems,
        [DefaultValue(true)] bool clearFirst = true)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_slicer",
            action.ToActionString(),
            () =>
            {
                var pivotCommands = new PivotTableCommands();
                var tableCommands = new TableCommands();

                return action switch
                {
                    // PivotTable slicers
                    SlicerAction.CreateSlicer => CreateSlicer(pivotCommands, sessionId, pivotTableName, fieldName, slicerName, destinationSheet, position),
                    SlicerAction.ListSlicers => ListSlicers(pivotCommands, sessionId, pivotTableName),
                    SlicerAction.SetSlicerSelection => SetSlicerSelection(pivotCommands, sessionId, slicerName, selectedItems, clearFirst),
                    SlicerAction.DeleteSlicer => DeleteSlicer(pivotCommands, sessionId, slicerName),
                    // Table slicers
                    SlicerAction.CreateTableSlicer => CreateTableSlicer(tableCommands, sessionId, tableName, columnName, slicerName, destinationSheet, position),
                    SlicerAction.ListTableSlicers => ListTableSlicers(tableCommands, sessionId, tableName),
                    SlicerAction.SetTableSlicerSelection => SetTableSlicerSelection(tableCommands, sessionId, slicerName, selectedItems, clearFirst),
                    SlicerAction.DeleteTableSlicer => DeleteTableSlicer(tableCommands, sessionId, slicerName),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    #region PivotTable Slicer Methods

    private static string CreateSlicer(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName,
        string? fieldName,
        string? slicerName,
        string? destinationSheet,
        string? position)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "create-slicer");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "create-slicer");
        if (string.IsNullOrWhiteSpace(slicerName))
            ExcelToolsBase.ThrowMissingParameter("slicerName", "create-slicer");
        if (string.IsNullOrWhiteSpace(destinationSheet))
            ExcelToolsBase.ThrowMissingParameter("destinationSheet", "create-slicer");
        if (string.IsNullOrWhiteSpace(position))
            ExcelToolsBase.ThrowMissingParameter("position", "create-slicer");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateSlicer(batch, pivotTableName!, fieldName!, slicerName!, destinationSheet!, position!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Name,
            result.Caption,
            result.FieldName,
            result.SheetName,
            result.Position,
            result.SelectedItems,
            result.AvailableItems,
            result.ConnectedPivotTables,
            result.SourceType,
            result.WorkflowHint,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string ListSlicers(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.ListSlicers(batch, pivotTableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Slicers,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string SetSlicerSelection(
        PivotTableCommands commands,
        string sessionId,
        string? slicerName,
        string? selectedItems,
        bool clearFirst)
    {
        if (string.IsNullOrWhiteSpace(slicerName))
            ExcelToolsBase.ThrowMissingParameter("slicerName", "set-slicer-selection");

        // Parse the selectedItems JSON array
        List<string> items = ParseSelectedItems(selectedItems);

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetSlicerSelection(batch, slicerName!, items, clearFirst));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Name,
            result.Caption,
            result.FieldName,
            result.SheetName,
            result.Position,
            result.SelectedItems,
            result.AvailableItems,
            result.ConnectedPivotTables,
            result.SourceType,
            result.WorkflowHint,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string DeleteSlicer(
        PivotTableCommands commands,
        string sessionId,
        string? slicerName)
    {
        if (string.IsNullOrWhiteSpace(slicerName))
            ExcelToolsBase.ThrowMissingParameter("slicerName", "delete-slicer");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.DeleteSlicer(batch, slicerName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
    }

    #endregion

    #region Table Slicer Methods

    private static string CreateTableSlicer(
        TableCommands commands,
        string sessionId,
        string? tableName,
        string? columnName,
        string? slicerName,
        string? destinationSheet,
        string? position)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter("tableName", "create-table-slicer");
        if (string.IsNullOrWhiteSpace(columnName))
            ExcelToolsBase.ThrowMissingParameter("columnName", "create-table-slicer");
        if (string.IsNullOrWhiteSpace(slicerName))
            ExcelToolsBase.ThrowMissingParameter("slicerName", "create-table-slicer");
        if (string.IsNullOrWhiteSpace(destinationSheet))
            ExcelToolsBase.ThrowMissingParameter("destinationSheet", "create-table-slicer");
        if (string.IsNullOrWhiteSpace(position))
            ExcelToolsBase.ThrowMissingParameter("position", "create-table-slicer");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateTableSlicer(batch, tableName!, columnName!, slicerName!, destinationSheet!, position!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Name,
            result.Caption,
            result.FieldName,
            result.SheetName,
            result.Position,
            result.SelectedItems,
            result.AvailableItems,
            result.ConnectedTable,
            result.SourceType,
            result.WorkflowHint,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string ListTableSlicers(
        TableCommands commands,
        string sessionId,
        string? tableName)
    {
        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.ListTableSlicers(batch, tableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Slicers,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string SetTableSlicerSelection(
        TableCommands commands,
        string sessionId,
        string? slicerName,
        string? selectedItems,
        bool clearFirst)
    {
        if (string.IsNullOrWhiteSpace(slicerName))
            ExcelToolsBase.ThrowMissingParameter("slicerName", "set-table-slicer-selection");

        // Parse the selectedItems JSON array
        List<string> items = ParseSelectedItems(selectedItems);

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetTableSlicerSelection(batch, slicerName!, items, clearFirst));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Name,
            result.Caption,
            result.FieldName,
            result.SheetName,
            result.Position,
            result.SelectedItems,
            result.AvailableItems,
            result.ConnectedTable,
            result.SourceType,
            result.WorkflowHint,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string DeleteTableSlicer(
        TableCommands commands,
        string sessionId,
        string? slicerName)
    {
        if (string.IsNullOrWhiteSpace(slicerName))
            ExcelToolsBase.ThrowMissingParameter("slicerName", "delete-table-slicer");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.DeleteTableSlicer(batch, slicerName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
    }

    #endregion

    #region Shared Helper Methods

    private static List<string> ParseSelectedItems(string? selectedItems)
    {
        if (string.IsNullOrWhiteSpace(selectedItems))
            return [];

        try
        {
            return JsonSerializer.Deserialize<List<string>>(selectedItems, JsonOptions) ?? [];
        }
        catch (JsonException)
        {
            // If parsing fails, treat as single item
            return [selectedItems!];
        }
    }

    #endregion
}
