using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Slicer operations
/// </summary>
[McpServerToolType]
public static partial class ExcelSlicerTool
{
    /// <summary>
    /// Slicer management: create, list, configure, and delete visual filtering controls.
    ///
    /// SLICERS provide visual filtering for PivotTables AND Tables:
    /// - Click items in the slicer to filter connected data
    /// - PivotTable slicers can filter multiple PivotTables
    /// - Table slicers filter a single Excel Table
    ///
    /// BEST PRACTICE - SLICER NAMING:
    /// When user does not specify a slicer name, auto-generate a descriptive name
    /// based on the field/column being filtered. Pattern: {FieldName}Slicer
    /// Examples: RegionSlicer, CategorySlicer, DepartmentSlicer, YearSlicer
    /// Do NOT ask the user to provide a slicer name - generate one automatically.
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
    [McpServerTool(Name = "excel_slicer", Title = "Excel Slicer Operations", Destructive = true)]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelSlicer(
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
        [DefaultValue(true)] bool clearFirst)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_slicer",
            ServiceRegistry.Slicer.ToActionString(action),
            () =>
            {
                return action switch
                {
                    // PivotTable slicers
                    SlicerAction.CreateSlicer => ForwardCreateSlicer(sessionId, pivotTableName, fieldName, slicerName, destinationSheet, position),
                    SlicerAction.ListSlicers => ForwardListSlicers(sessionId, pivotTableName),
                    SlicerAction.SetSlicerSelection => ForwardSetSlicerSelection(sessionId, slicerName, selectedItems, clearFirst),
                    SlicerAction.DeleteSlicer => ForwardDeleteSlicer(sessionId, slicerName),
                    // Table slicers
                    SlicerAction.CreateTableSlicer => ForwardCreateTableSlicer(sessionId, tableName, columnName, slicerName, destinationSheet, position),
                    SlicerAction.ListTableSlicers => ForwardListTableSlicers(sessionId, tableName),
                    SlicerAction.SetTableSlicerSelection => ForwardSetTableSlicerSelection(sessionId, slicerName, selectedItems, clearFirst),
                    SlicerAction.DeleteTableSlicer => ForwardDeleteTableSlicer(sessionId, slicerName),
                    _ => throw new ArgumentException($"Unknown action: {action} ({ServiceRegistry.Slicer.ToActionString(action)})", nameof(action))
                };
            });
    }

    // === PIVOTTABLE SLICER SERVICE FORWARDING METHODS ===

    private static string ForwardCreateSlicer(
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
        if (string.IsNullOrWhiteSpace(destinationSheet))
            ExcelToolsBase.ThrowMissingParameter("destinationSheet", "create-slicer");
        if (string.IsNullOrWhiteSpace(position))
            ExcelToolsBase.ThrowMissingParameter("position", "create-slicer");

        // Auto-generate slicer name from field name if not provided
        var effectiveSlicerName = string.IsNullOrWhiteSpace(slicerName)
            ? $"{fieldName}Slicer"
            : slicerName;

        // Parse position string to left/top coordinates (position like "E1" or "100,100")
        var (left, top) = ParsePosition(position!);

        return ExcelToolsBase.ForwardToService("slicer.create-slicer", sessionId, new
        {
            pivotTableName,
            sourceFieldName = fieldName,
            slicerName = effectiveSlicerName,
            destinationSheet,
            left,
            top
        });
    }

    private static string ForwardListSlicers(string sessionId, string? pivotTableName)
    {
        return ExcelToolsBase.ForwardToService("slicer.list-slicers", sessionId, new { pivotTableName });
    }

    private static string ForwardSetSlicerSelection(string sessionId, string? slicerName, string? selectedItems, bool clearFirst)
    {
        if (string.IsNullOrWhiteSpace(slicerName))
            ExcelToolsBase.ThrowMissingParameter("slicerName", "set-slicer-selection");

        // Parse the selectedItems JSON array to comma-separated string for service
        var itemsCsv = ParseSelectedItemsToCsv(selectedItems);

        return ExcelToolsBase.ForwardToService("slicer.set-slicer-selection", sessionId, new
        {
            slicerName,
            selectedItems = itemsCsv,
            multiSelect = !clearFirst
        });
    }

    private static string ForwardDeleteSlicer(string sessionId, string? slicerName)
    {
        if (string.IsNullOrWhiteSpace(slicerName))
            ExcelToolsBase.ThrowMissingParameter("slicerName", "delete-slicer");

        return ExcelToolsBase.ForwardToService("slicer.delete-slicer", sessionId, new { slicerName });
    }

    // === TABLE SLICER SERVICE FORWARDING METHODS ===

    private static string ForwardCreateTableSlicer(
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
        if (string.IsNullOrWhiteSpace(destinationSheet))
            ExcelToolsBase.ThrowMissingParameter("destinationSheet", "create-table-slicer");
        if (string.IsNullOrWhiteSpace(position))
            ExcelToolsBase.ThrowMissingParameter("position", "create-table-slicer");

        // Auto-generate slicer name from column name if not provided
        var effectiveSlicerName = string.IsNullOrWhiteSpace(slicerName)
            ? $"{columnName}Slicer"
            : slicerName;

        // Parse position string to left/top coordinates
        var (left, top) = ParsePosition(position!);

        return ExcelToolsBase.ForwardToService("slicer.create-table-slicer", sessionId, new
        {
            tableName,
            columnName,
            slicerName = effectiveSlicerName,
            destinationSheet,
            left,
            top
        });
    }

    private static string ForwardListTableSlicers(string sessionId, string? tableName)
    {
        return ExcelToolsBase.ForwardToService("slicer.list-table-slicers", sessionId, new { tableName });
    }

    private static string ForwardSetTableSlicerSelection(string sessionId, string? slicerName, string? selectedItems, bool clearFirst)
    {
        if (string.IsNullOrWhiteSpace(slicerName))
            ExcelToolsBase.ThrowMissingParameter("slicerName", "set-table-slicer-selection");

        // Parse the selectedItems JSON array to comma-separated string for service
        var itemsCsv = ParseSelectedItemsToCsv(selectedItems);

        return ExcelToolsBase.ForwardToService("slicer.set-table-slicer-selection", sessionId, new
        {
            slicerName,
            selectedItems = itemsCsv,
            multiSelect = !clearFirst
        });
    }

    private static string ForwardDeleteTableSlicer(string sessionId, string? slicerName)
    {
        if (string.IsNullOrWhiteSpace(slicerName))
            ExcelToolsBase.ThrowMissingParameter("slicerName", "delete-table-slicer");

        return ExcelToolsBase.ForwardToService("slicer.delete-table-slicer", sessionId, new { slicerName });
    }

    // === HELPER METHODS ===

    /// <summary>
    /// Parse position string to left/top coordinates.
    /// Supports "100,100" format or defaults to 100,100 for cell addresses.
    /// </summary>
    private static (double left, double top) ParsePosition(string position)
    {
        if (position.Contains(','))
        {
            var parts = position.Split(',');
            if (parts.Length >= 2 &&
                double.TryParse(parts[0].Trim(), out var left) &&
                double.TryParse(parts[1].Trim(), out var top))
            {
                return (left, top);
            }
        }

        // Default position for cell addresses or invalid formats
        return (100, 100);
    }

    /// <summary>
    /// Parse JSON array of selected items to comma-separated string for service.
    /// </summary>
    private static string? ParseSelectedItemsToCsv(string? selectedItems)
    {
        if (string.IsNullOrWhiteSpace(selectedItems))
            return null;

        try
        {
            var items = JsonSerializer.Deserialize<List<string>>(selectedItems, ExcelToolsBase.JsonOptions);
            return items != null ? string.Join(",", items) : null;
        }
        catch (JsonException)
        {
            // If parsing fails, treat as single item
            return selectedItems;
        }
    }
}




