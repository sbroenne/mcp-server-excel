using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel PivotTable operations
/// </summary>
[McpServerToolType]
public static partial class ExcelPivotTableTool
{
    private static readonly JsonSerializerOptions JsonOptions = ExcelToolsBase.JsonOptions;

    /// <summary>
    /// PivotTable lifecycle management: create from various sources, list, read details, refresh, and delete.
    ///
    /// BEST PRACTICE: Use 'list' before creating. Prefer 'refresh' or field modifications over delete+recreate.
    /// Delete+recreate loses field configurations, filters, sorting, and custom layouts.
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
    [McpServerTool(Name = "excel_pivottable", Title = "Excel PivotTable Operations")]
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
            action.ToActionString(),
            () =>
            {
                var commands = new PivotTableCommands();

                return action switch
                {
                    PivotTableAction.List => List(commands, sessionId),
                    PivotTableAction.Read => Read(commands, sessionId, pivotTableName),
                    PivotTableAction.CreateFromRange => CreateFromRange(commands, sessionId, sourceSheetName, sourceRangeAddress, destinationSheetName, destinationCellAddress, pivotTableName, layoutStyle),
                    PivotTableAction.CreateFromTable => CreateFromTable(commands, sessionId, sourceTableName, destinationSheetName, destinationCellAddress, pivotTableName, layoutStyle),
                    PivotTableAction.CreateFromDataModel => CreateFromDataModel(commands, sessionId, dataModelTableName, destinationSheetName, destinationCellAddress, pivotTableName, layoutStyle),
                    PivotTableAction.Delete => Delete(commands, sessionId, pivotTableName),
                    PivotTableAction.Refresh => Refresh(commands, sessionId, pivotTableName),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string List(
        PivotTableCommands commands,
        string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTables,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string Read(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "read");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Read(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTable,
            result.Fields,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string CreateFromRange(
        PivotTableCommands commands,
        string sessionId,
        string? sourceSheetName,
        string? sourceRangeAddress,
        string? destinationSheetName,
        string? destinationCellAddress,
        string? pivotTableName,
        int? layoutStyle)
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch =>
            {
                var createResult = commands.CreateFromRange(batch, sourceSheetName!, sourceRangeAddress!,
                    destinationSheetName!, destinationCellAddress!, pivotTableName!);

                // Apply layout if specified and creation succeeded
                if (createResult.Success && layoutStyle.HasValue)
                {
                    commands.SetLayout(batch, createResult.PivotTableName!, layoutStyle.Value);
                }

                return createResult;
            });

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.SheetName,
            result.Range,
            result.SourceData,
            result.SourceRowCount,
            result.AvailableFields,
            result.ErrorMessage,
            LayoutApplied = result.Success && layoutStyle.HasValue ? layoutStyle : null
        }, JsonOptions);
    }

    private static string CreateFromTable(
        PivotTableCommands commands,
        string sessionId,
        string? sourceTableName,
        string? destinationSheetName,
        string? destinationCellAddress,
        string? pivotTableName,
        int? layoutStyle)
    {
        if (string.IsNullOrWhiteSpace(sourceTableName))
            ExcelToolsBase.ThrowMissingParameter("sourceTableName", "create-from-table");
        if (string.IsNullOrWhiteSpace(destinationSheetName))
            ExcelToolsBase.ThrowMissingParameter("destinationSheetName", "create-from-table");
        if (string.IsNullOrWhiteSpace(destinationCellAddress))
            ExcelToolsBase.ThrowMissingParameter("destinationCellAddress", "create-from-table");
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "create-from-table");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch =>
            {
                var createResult = commands.CreateFromTable(batch, sourceTableName!,
                    destinationSheetName!, destinationCellAddress!, pivotTableName!);

                // Apply layout if specified and creation succeeded
                if (createResult.Success && layoutStyle.HasValue)
                {
                    commands.SetLayout(batch, createResult.PivotTableName!, layoutStyle.Value);
                }

                return createResult;
            });

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.SheetName,
            result.Range,
            result.SourceData,
            result.SourceRowCount,
            result.AvailableFields,
            result.ErrorMessage,
            LayoutApplied = result.Success && layoutStyle.HasValue ? layoutStyle : null
        }, JsonOptions);
    }

    private static string CreateFromDataModel(
        PivotTableCommands commands,
        string sessionId,
        string? dataModelTableName,
        string? destinationSheetName,
        string? destinationCellAddress,
        string? pivotTableName,
        int? layoutStyle)
    {
        if (string.IsNullOrWhiteSpace(dataModelTableName))
            ExcelToolsBase.ThrowMissingParameter("dataModelTableName", "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(destinationSheetName))
            ExcelToolsBase.ThrowMissingParameter("destinationSheetName", "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(destinationCellAddress))
            ExcelToolsBase.ThrowMissingParameter("destinationCellAddress", "create-from-datamodel");
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "create-from-datamodel");

        PivotTableCreateResult result;
        int? appliedLayout = null;

        try
        {
            result = ExcelToolsBase.WithSession(sessionId,
                batch =>
                {
                    var createResult = commands.CreateFromDataModel(batch, dataModelTableName!,
                        destinationSheetName!, destinationCellAddress!, pivotTableName!);

                    // Apply layout if specified and creation succeeded
                    if (createResult.Success && layoutStyle.HasValue)
                    {
                        commands.SetLayout(batch, createResult.PivotTableName!, layoutStyle.Value);
                    }

                    return createResult;
                });

            if (result.Success && layoutStyle.HasValue)
            {
                appliedLayout = layoutStyle;
            }
        }
        catch (TimeoutException ex)
        {
            result = new PivotTableCreateResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                PivotTableName = pivotTableName!,
                SheetName = destinationSheetName!,
                Range = string.Empty,
                SourceData = dataModelTableName!,
                SourceRowCount = 0,
                AvailableFields = []
            };
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.SheetName,
            result.Range,
            result.SourceData,
            result.SourceRowCount,
            result.AvailableFields,
            result.ErrorMessage,
            isError = !result.Success,
            LayoutApplied = appliedLayout
        }, JsonOptions);
    }

    private static string Delete(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "delete");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Delete(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string Refresh(
        PivotTableCommands commands,
        string sessionId,
        string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "refresh");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Refresh(batch, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.RefreshTime,
            result.SourceRecordCount,
            result.PreviousRecordCount,
            result.StructureChanged,
            result.NewFields,
            result.RemovedFields,
            result.ErrorMessage
        }, JsonOptions);
    }
}
