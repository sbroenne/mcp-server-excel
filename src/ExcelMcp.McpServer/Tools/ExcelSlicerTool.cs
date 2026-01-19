using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;

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
    /// SLICERS provide visual filtering for PivotTables:
    /// - Click items in the slicer to filter connected PivotTables
    /// - A single slicer can filter multiple PivotTables
    /// - Each slicer is backed by a SlicerCache that manages the field data
    ///
    /// CREATE: Specify pivotTableName, fieldName (from PivotTable fields), slicerName,
    /// destinationSheet, and position (cell address like 'E1').
    ///
    /// LIST: Returns all slicers. Optionally filter by pivotTableName.
    ///
    /// SET-SLICER-SELECTION: Pass selectedItems as JSON array of strings.
    /// Empty array clears filter (shows all items). Set clearFirst=false to add
    /// to existing selection instead of replacing.
    ///
    /// DELETE: Removes the visual slicer. SlicerCache is auto-deleted if no more
    /// slicers use it.
    ///
    /// RELATED TOOLS:
    /// - excel_pivottable: Create PivotTables (required before creating slicers)
    /// - excel_pivottable_field: Manage PivotTable fields (use list-fields to see available field names)
    /// </summary>
    /// <param name="action">The slicer operation to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action</param>
    /// <param name="pivotTableName">Name of the PivotTable - required for create-slicer, optional for list-slicers (filters results)</param>
    /// <param name="fieldName">Field name from the PivotTable to create slicer for (required for create-slicer)</param>
    /// <param name="slicerName">Name for the slicer - required for create-slicer, set-slicer-selection, delete-slicer</param>
    /// <param name="destinationSheet">Worksheet where slicer will be placed (required for create-slicer)</param>
    /// <param name="position">Cell address for top-left corner of slicer, e.g., 'E1' (required for create-slicer)</param>
    /// <param name="selectedItems">JSON array of item names to select, e.g., '["North","South"]'. Empty array clears filter.</param>
    /// <param name="clearFirst">If true (default), replaces selection. If false, adds to existing selection.</param>
    [McpServerTool(Name = "excel_slicer", Title = "Excel Slicer Operations")]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static string ExcelSlicer(
        SlicerAction action,
        string sessionId,
        [DefaultValue(null)] string? pivotTableName,
        [DefaultValue(null)] string? fieldName,
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
                var commands = new PivotTableCommands();

                return action switch
                {
                    SlicerAction.CreateSlicer => CreateSlicer(commands, sessionId, pivotTableName, fieldName, slicerName, destinationSheet, position),
                    SlicerAction.ListSlicers => ListSlicers(commands, sessionId, pivotTableName),
                    SlicerAction.SetSlicerSelection => SetSlicerSelection(commands, sessionId, slicerName, selectedItems, clearFirst),
                    SlicerAction.DeleteSlicer => DeleteSlicer(commands, sessionId, slicerName),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

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
        List<string> items = [];
        if (!string.IsNullOrWhiteSpace(selectedItems))
        {
            try
            {
                items = JsonSerializer.Deserialize<List<string>>(selectedItems, JsonOptions) ?? [];
            }
            catch (JsonException)
            {
                // If parsing fails, treat as single item
                items = [selectedItems!];
            }
        }

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
}
