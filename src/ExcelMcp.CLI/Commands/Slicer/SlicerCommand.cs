using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.Slicer;

/// <summary>
/// CLI command for slicer operations on PivotTables and Excel Tables.
/// </summary>
internal sealed class SlicerCommand : SessionCommandBase<SlicerCommand.Settings>
{
    private readonly IPivotTableCommands _pivotCommands;
    private readonly ITableCommands _tableCommands;

    public SlicerCommand(
        ISessionService sessionService,
        IPivotTableCommands pivotCommands,
        ITableCommands tableCommands,
        ICliConsole console)
        : base(sessionService, console)
    {
        _pivotCommands = pivotCommands ?? throw new ArgumentNullException(nameof(pivotCommands));
        _tableCommands = tableCommands ?? throw new ArgumentNullException(nameof(tableCommands));
    }

    protected override string CommandName => "slicer";

    protected override int ExecuteAction(
        CommandContext context,
        Settings settings,
        IExcelBatch batch,
        string action,
        CancellationToken cancellationToken)
    {
        return action switch
        {
            // PivotTable slicer actions
            "create-slicer" => ExecuteCreateSlicer(batch, settings),
            "list-slicers" => WriteResult(_pivotCommands.ListSlicers(batch, settings.PivotTableName)),
            "set-slicer-selection" => ExecuteSetSlicerSelection(batch, settings),
            "delete-slicer" => ExecuteDeleteSlicer(batch, settings),
            // Table slicer actions
            "create-table-slicer" => ExecuteCreateTableSlicer(batch, settings),
            "list-table-slicers" => WriteResult(_tableCommands.ListTableSlicers(batch, settings.TableName)),
            "set-table-slicer-selection" => ExecuteSetTableSlicerSelection(batch, settings),
            "delete-table-slicer" => ExecuteDeleteTableSlicer(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    #region PivotTable Slicer Actions

    private int ExecuteCreateSlicer(IExcelBatch batch, Settings settings)
    {
        if (!RequireParameters("create-slicer",
            (settings.PivotTableName, "pivot-name"),
            (settings.FieldName, "field-name"),
            (settings.DestinationSheet, "destination-sheet"),
            (settings.Position, "position")))
            return ExitCodes.MissingParameter;

        // Auto-generate slicer name if not provided
        var slicerName = string.IsNullOrWhiteSpace(settings.SlicerName)
            ? $"{settings.FieldName}Slicer"
            : settings.SlicerName;

        return WriteResult(_pivotCommands.CreateSlicer(
            batch,
            settings.PivotTableName!,
            settings.FieldName!,
            slicerName,
            settings.DestinationSheet!,
            settings.Position!));
    }

    private int ExecuteSetSlicerSelection(IExcelBatch batch, Settings settings)
    {
        if (!RequireParameter(settings.SlicerName, "slicer-name", "set-slicer-selection"))
            return ExitCodes.MissingParameter;

        var items = ParseSelectedItems(settings.SelectedItems);

        return WriteResult(_pivotCommands.SetSlicerSelection(
            batch,
            settings.SlicerName!,
            items,
            settings.ClearFirst));
    }

    private int ExecuteDeleteSlicer(IExcelBatch batch, Settings settings)
    {
        if (!RequireParameter(settings.SlicerName, "slicer-name", "delete-slicer"))
            return ExitCodes.MissingParameter;

        return WriteResult(_pivotCommands.DeleteSlicer(batch, settings.SlicerName!));
    }

    #endregion

    #region Table Slicer Actions

    private int ExecuteCreateTableSlicer(IExcelBatch batch, Settings settings)
    {
        if (!RequireParameters("create-table-slicer",
            (settings.TableName, "table-name"),
            (settings.ColumnName, "column-name"),
            (settings.DestinationSheet, "destination-sheet"),
            (settings.Position, "position")))
            return ExitCodes.MissingParameter;

        // Auto-generate slicer name if not provided
        var slicerName = string.IsNullOrWhiteSpace(settings.SlicerName)
            ? $"{settings.ColumnName}Slicer"
            : settings.SlicerName;

        return WriteResult(_tableCommands.CreateTableSlicer(
            batch,
            settings.TableName!,
            settings.ColumnName!,
            slicerName,
            settings.DestinationSheet!,
            settings.Position!));
    }

    private int ExecuteSetTableSlicerSelection(IExcelBatch batch, Settings settings)
    {
        if (!RequireParameter(settings.SlicerName, "slicer-name", "set-table-slicer-selection"))
            return ExitCodes.MissingParameter;

        var items = ParseSelectedItems(settings.SelectedItems);

        return WriteResult(_tableCommands.SetTableSlicerSelection(
            batch,
            settings.SlicerName!,
            items,
            settings.ClearFirst));
    }

    private int ExecuteDeleteTableSlicer(IExcelBatch batch, Settings settings)
    {
        if (!RequireParameter(settings.SlicerName, "slicer-name", "delete-table-slicer"))
            return ExitCodes.MissingParameter;

        return WriteResult(_tableCommands.DeleteTableSlicer(batch, settings.SlicerName!));
    }

    #endregion

    #region Helpers

    private static List<string> ParseSelectedItems(string? selectedItems)
    {
        if (string.IsNullOrWhiteSpace(selectedItems))
            return [];

        // Try parsing as JSON array first
        try
        {
            return JsonSerializer.Deserialize<List<string>>(selectedItems) ?? [];
        }
        catch (JsonException)
        {
            // Fall back to comma-separated values
            return selectedItems
                .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                .ToList();
        }
    }

    #endregion

    internal sealed class Settings : SessionSettings
    {
        // PivotTable slicer options
        [CommandOption("--pivot-name <NAME>")]
        public string? PivotTableName { get; init; }

        [CommandOption("--field-name <NAME>")]
        public string? FieldName { get; init; }

        // Table slicer options
        [CommandOption("--table-name <NAME>")]
        public string? TableName { get; init; }

        [CommandOption("--column-name <NAME>")]
        public string? ColumnName { get; init; }

        // Common options
        [CommandOption("--slicer-name <NAME>")]
        public string? SlicerName { get; init; }

        [CommandOption("--destination-sheet <SHEET>")]
        public string? DestinationSheet { get; init; }

        [CommandOption("--position <CELL>")]
        public string? Position { get; init; }

        [CommandOption("--selected-items <JSON_OR_CSV>")]
        public string? SelectedItems { get; init; }

        [CommandOption("--clear-first")]
        public bool ClearFirst { get; init; } = true;
    }
}
