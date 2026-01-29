using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.Slicer;

/// <summary>
/// CLI command for slicer operations on PivotTables and Excel Tables.
/// </summary>
internal sealed class SlicerCommand : Command<SlicerCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly IPivotTableCommands _pivotCommands;
    private readonly ITableCommands _tableCommands;
    private readonly ICliConsole _console;

    public SlicerCommand(
        ISessionService sessionService,
        IPivotTableCommands pivotCommands,
        ITableCommands tableCommands,
        ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _pivotCommands = pivotCommands ?? throw new ArgumentNullException(nameof(pivotCommands));
        _tableCommands = tableCommands ?? throw new ArgumentNullException(nameof(tableCommands));
        _console = console ?? throw new ArgumentNullException(nameof(console));
    }

    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            _console.WriteError("Session ID is required. Use 'session open' first.");
            return -1;
        }

        var action = settings.Action?.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(action))
        {
            _console.WriteError("Action is required.");
            return -1;
        }

        var batch = _sessionService.GetBatch(settings.SessionId);

        return action switch
        {
            // PivotTable slicer actions
            "create-slicer" => ExecuteCreateSlicer(batch, settings),
            "list-slicers" => ExecuteListSlicers(batch, settings),
            "set-slicer-selection" => ExecuteSetSlicerSelection(batch, settings),
            "delete-slicer" => ExecuteDeleteSlicer(batch, settings),
            // Table slicer actions
            "create-table-slicer" => ExecuteCreateTableSlicer(batch, settings),
            "list-table-slicers" => ExecuteListTableSlicers(batch, settings),
            "set-table-slicer-selection" => ExecuteSetTableSlicerSelection(batch, settings),
            "delete-table-slicer" => ExecuteDeleteTableSlicer(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    #region PivotTable Slicer Actions

    private int ExecuteCreateSlicer(ComInterop.Session.IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for create-slicer.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.FieldName))
        {
            _console.WriteError("--field-name is required for create-slicer.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.DestinationSheet))
        {
            _console.WriteError("--destination-sheet is required for create-slicer.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.Position))
        {
            _console.WriteError("--position is required for create-slicer (e.g., 'E1').");
            return -1;
        }

        // Auto-generate slicer name if not provided
        var slicerName = string.IsNullOrWhiteSpace(settings.SlicerName)
            ? $"{settings.FieldName}Slicer"
            : settings.SlicerName;

        return WriteResult(_pivotCommands.CreateSlicer(
            batch,
            settings.PivotTableName,
            settings.FieldName,
            slicerName,
            settings.DestinationSheet,
            settings.Position));
    }

    private int ExecuteListSlicers(ComInterop.Session.IExcelBatch batch, Settings settings)
    {
        return WriteResult(_pivotCommands.ListSlicers(batch, settings.PivotTableName));
    }

    private int ExecuteSetSlicerSelection(ComInterop.Session.IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SlicerName))
        {
            _console.WriteError("--slicer-name is required for set-slicer-selection.");
            return -1;
        }

        var items = ParseSelectedItems(settings.SelectedItems);

        return WriteResult(_pivotCommands.SetSlicerSelection(
            batch,
            settings.SlicerName,
            items,
            settings.ClearFirst));
    }

    private int ExecuteDeleteSlicer(ComInterop.Session.IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SlicerName))
        {
            _console.WriteError("--slicer-name is required for delete-slicer.");
            return -1;
        }

        return WriteResult(_pivotCommands.DeleteSlicer(batch, settings.SlicerName));
    }

    #endregion

    #region Table Slicer Actions

    private int ExecuteCreateTableSlicer(ComInterop.Session.IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName))
        {
            _console.WriteError("--table-name is required for create-table-slicer.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.ColumnName))
        {
            _console.WriteError("--column-name is required for create-table-slicer.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.DestinationSheet))
        {
            _console.WriteError("--destination-sheet is required for create-table-slicer.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.Position))
        {
            _console.WriteError("--position is required for create-table-slicer (e.g., 'E1').");
            return -1;
        }

        // Auto-generate slicer name if not provided
        var slicerName = string.IsNullOrWhiteSpace(settings.SlicerName)
            ? $"{settings.ColumnName}Slicer"
            : settings.SlicerName;

        return WriteResult(_tableCommands.CreateTableSlicer(
            batch,
            settings.TableName,
            settings.ColumnName,
            slicerName,
            settings.DestinationSheet,
            settings.Position));
    }

    private int ExecuteListTableSlicers(ComInterop.Session.IExcelBatch batch, Settings settings)
    {
        return WriteResult(_tableCommands.ListTableSlicers(batch, settings.TableName));
    }

    private int ExecuteSetTableSlicerSelection(ComInterop.Session.IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SlicerName))
        {
            _console.WriteError("--slicer-name is required for set-table-slicer-selection.");
            return -1;
        }

        var items = ParseSelectedItems(settings.SelectedItems);

        return WriteResult(_tableCommands.SetTableSlicerSelection(
            batch,
            settings.SlicerName,
            items,
            settings.ClearFirst));
    }

    private int ExecuteDeleteTableSlicer(ComInterop.Session.IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SlicerName))
        {
            _console.WriteError("--slicer-name is required for delete-table-slicer.");
            return -1;
        }

        return WriteResult(_tableCommands.DeleteTableSlicer(batch, settings.SlicerName));
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

    private int WriteResult(ResultBase result)
    {
        _console.WriteJson(result);
        return result.Success ? 0 : -1;
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown slicer action '{action}'.");
        return -1;
    }

    #endregion

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

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
