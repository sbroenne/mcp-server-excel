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
    /// LAYOUT: Use excel_pivottable_calc set-layout action (0=Compact, 1=Tabular, 2=Outline).
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
    /// - excel_pivottable_calc: Calculated fields, layout options, subtotals
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
        [DefaultValue(null)] string? destinationCellAddress)
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
                    PivotTableAction.CreateFromRange => CreateFromRange(commands, sessionId, sourceSheetName, sourceRangeAddress, destinationSheetName, destinationCellAddress, pivotTableName),
                    PivotTableAction.CreateFromTable => CreateFromTable(commands, sessionId, sourceTableName, destinationSheetName, destinationCellAddress, pivotTableName),
                    PivotTableAction.CreateFromDataModel => CreateFromDataModel(commands, sessionId, dataModelTableName, destinationSheetName, destinationCellAddress, pivotTableName),
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateFromRange(batch, sourceSheetName!, sourceRangeAddress!,
                destinationSheetName!, destinationCellAddress!, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.SheetName,
            result.Range,
            result.SourceData,
            result.SourceRowCount,
            result.AvailableFields,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string CreateFromTable(
        PivotTableCommands commands,
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateFromTable(batch, sourceTableName!,
                destinationSheetName!, destinationCellAddress!, pivotTableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.PivotTableName,
            result.SheetName,
            result.Range,
            result.SourceData,
            result.SourceRowCount,
            result.AvailableFields,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string CreateFromDataModel(
        PivotTableCommands commands,
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

        PivotTableCreateResult result;

        try
        {
            result = ExcelToolsBase.WithSession(sessionId,
                batch => commands.CreateFromDataModel(batch, dataModelTableName!,
                    destinationSheetName!, destinationCellAddress!, pivotTableName!));
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
            isError = !result.Success
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
