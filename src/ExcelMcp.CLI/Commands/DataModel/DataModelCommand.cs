using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.DataModel;

internal sealed class DataModelCommand : Command<DataModelCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly IDataModelCommands _dataModelCommands;
    private readonly ICliConsole _console;

    public DataModelCommand(ISessionService sessionService, IDataModelCommands dataModelCommands, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _dataModelCommands = dataModelCommands ?? throw new ArgumentNullException(nameof(dataModelCommands));
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
            "list-tables" => WriteResult(_dataModelCommands.ListTables(batch)),
            "list-columns" => ExecuteListColumns(batch, settings),
            "read-table" or "get-table" => ExecuteReadTable(batch, settings),
            "read-info" or "info" => WriteResult(_dataModelCommands.ReadInfo(batch)),
            "list-measures" => ExecuteListMeasures(batch, settings),
            "read" or "read-measure" or "get-measure" => ExecuteReadMeasure(batch, settings),
            "create-measure" => ExecuteCreateMeasure(batch, settings),
            "update-measure" => ExecuteUpdateMeasure(batch, settings),
            "delete-measure" => ExecuteDeleteMeasure(batch, settings),
            "rename-table" => ExecuteRenameTable(batch, settings),
            "delete-table" => ExecuteDeleteTable(batch, settings),
            "list-relationships" => WriteResult(_dataModelCommands.ListRelationships(batch)),
            "read-relationship" or "get-relationship" => ExecuteReadRelationship(batch, settings),
            "create-relationship" => ExecuteCreateRelationship(batch, settings),
            "update-relationship" => ExecuteUpdateRelationship(batch, settings),
            "delete-relationship" => ExecuteDeleteRelationship(batch, settings),
            "refresh" => ExecuteRefresh(batch, settings),
            "evaluate" => ExecuteEvaluate(batch, settings),
            "execute-dmv" => ExecuteDmv(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    private int ExecuteListColumns(IExcelBatch batch, Settings settings)
    {
        if (!TryGetTable(settings, out var table))
        {
            return -1;
        }

        return WriteResult(_dataModelCommands.ListColumns(batch, table));
    }

    private int ExecuteReadTable(IExcelBatch batch, Settings settings)
    {
        if (!TryGetTable(settings, out var table))
        {
            return -1;
        }

        return WriteResult(_dataModelCommands.ReadTable(batch, table));
    }

    private int ExecuteRenameTable(IExcelBatch batch, Settings settings)
    {
        if (!TryGetTable(settings, out var table))
        {
            return -1;
        }

        var newName = settings.NewName?.Trim();
        if (string.IsNullOrWhiteSpace(newName))
        {
            _console.WriteError("--new-name is required for rename-table.");
            return -1;
        }

        return WriteResult(_dataModelCommands.RenameTable(batch, table, newName));
    }

    private int ExecuteDeleteTable(IExcelBatch batch, Settings settings)
    {
        if (!TryGetTable(settings, out var table))
        {
            return -1;
        }

        try
        {
            _dataModelCommands.DeleteTable(batch, table);
            _console.WriteJson(new { success = true, message = $"Table '{table}' deleted successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Error deleting table: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteListMeasures(IExcelBatch batch, Settings settings)
    {
        return WriteResult(_dataModelCommands.ListMeasures(batch, settings.TableName));
    }

    private int ExecuteReadMeasure(IExcelBatch batch, Settings settings)
    {
        if (!TryGetMeasure(settings, out var measure))
        {
            return -1;
        }

        return WriteResult(_dataModelCommands.Read(batch, measure));
    }

    private int ExecuteCreateMeasure(IExcelBatch batch, Settings settings)
    {
        if (!TryGetTable(settings, out var table) ||
            !TryGetMeasure(settings, out var measure) ||
            string.IsNullOrWhiteSpace(settings.DaxFormula))
        {
            if (string.IsNullOrWhiteSpace(settings.DaxFormula))
            {
                _console.WriteError("--dax is required for create-measure.");
            }
            return -1;
        }

        try
        {
            _dataModelCommands.CreateMeasure(
                batch,
                table,
                measure,
                settings.DaxFormula!,
                settings.FormatType,
                settings.Description);

            _console.WriteJson(new { success = true, message = $"Measure '{measure}' created successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Error creating measure: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteUpdateMeasure(IExcelBatch batch, Settings settings)
    {
        if (!TryGetMeasure(settings, out var measure))
        {
            return -1;
        }

        if (settings.DaxFormula is null && settings.FormatType is null && settings.Description is null)
        {
            _console.WriteError("Provide at least one of --dax, --format-type, or --description for update-measure.");
            return -1;
        }

        try
        {
            _dataModelCommands.UpdateMeasure(batch, measure, settings.DaxFormula, settings.FormatType, settings.Description);
            _console.WriteJson(new { success = true, message = $"Measure '{measure}' updated successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Error updating measure: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteDeleteMeasure(IExcelBatch batch, Settings settings)
    {
        if (!TryGetMeasure(settings, out var measure))
        {
            return -1;
        }

        try
        {
            _dataModelCommands.DeleteMeasure(batch, measure);
            _console.WriteJson(new { success = true, message = $"Measure '{measure}' deleted successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Error deleting measure: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteReadRelationship(IExcelBatch batch, Settings settings)
    {
        if (!TryGetRelationshipEndpoints(settings, out var fromTable, out var fromColumn, out var toTable, out var toColumn))
        {
            return -1;
        }

        return WriteResult(_dataModelCommands.ReadRelationship(batch, fromTable, fromColumn, toTable, toColumn));
    }

    private int ExecuteCreateRelationship(IExcelBatch batch, Settings settings)
    {
        if (!TryGetRelationshipEndpoints(settings, out var fromTable, out var fromColumn, out var toTable, out var toColumn))
        {
            return -1;
        }

        try
        {
            var active = settings.Active ?? true;
            _dataModelCommands.CreateRelationship(batch, fromTable, fromColumn, toTable, toColumn, active);
            _console.WriteJson(new { success = true, message = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} created successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Error creating relationship: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteUpdateRelationship(IExcelBatch batch, Settings settings)
    {
        if (!TryGetRelationshipEndpoints(settings, out var fromTable, out var fromColumn, out var toTable, out var toColumn))
        {
            return -1;
        }

        if (settings.Active is null)
        {
            _console.WriteError("--active is required for update-relationship.");
            return -1;
        }

        try
        {
            _dataModelCommands.UpdateRelationship(batch, fromTable, fromColumn, toTable, toColumn, settings.Active.Value);
            _console.WriteJson(new { success = true, message = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} updated successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Error updating relationship: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteDeleteRelationship(IExcelBatch batch, Settings settings)
    {
        if (!TryGetRelationshipEndpoints(settings, out var fromTable, out var fromColumn, out var toTable, out var toColumn))
        {
            return -1;
        }

        try
        {
            _dataModelCommands.DeleteRelationship(batch, fromTable, fromColumn, toTable, toColumn);
            _console.WriteJson(new { success = true, message = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} deleted successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Error deleting relationship: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteRefresh(IExcelBatch batch, Settings settings)
    {
        try
        {
            TimeSpan? timeout = settings.TimeoutSeconds.HasValue ? TimeSpan.FromSeconds(settings.TimeoutSeconds.Value) : null;
            _dataModelCommands.Refresh(batch, settings.TableName, timeout);
            _console.WriteJson(new { success = true, message = "Data Model refreshed successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Error refreshing Data Model: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteEvaluate(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.DaxQuery))
        {
            _console.WriteError("--dax-query is required for evaluate.");
            return -1;
        }

        return WriteResult(_dataModelCommands.Evaluate(batch, settings.DaxQuery));
    }

    private int ExecuteDmv(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.DmvQuery))
        {
            _console.WriteError("--dmv-query is required for execute-dmv.");
            return -1;
        }

        return WriteResult(_dataModelCommands.ExecuteDmv(batch, settings.DmvQuery));
    }

    private bool TryGetTable(Settings settings, out string tableName)
    {
        tableName = settings.TableName?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(tableName))
        {
            _console.WriteError("--table is required for this action.");
            return false;
        }

        return true;
    }

    private bool TryGetMeasure(Settings settings, out string measureName)
    {
        measureName = settings.MeasureName?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(measureName))
        {
            _console.WriteError("--measure is required for this action.");
            return false;
        }

        return true;
    }

    private bool TryGetRelationshipEndpoints(Settings settings, out string fromTable, out string fromColumn, out string toTable, out string toColumn)
    {
        fromTable = settings.FromTable?.Trim() ?? string.Empty;
        fromColumn = settings.FromColumn?.Trim() ?? string.Empty;
        toTable = settings.ToTable?.Trim() ?? string.Empty;
        toColumn = settings.ToColumn?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(fromTable) ||
            string.IsNullOrWhiteSpace(fromColumn) ||
            string.IsNullOrWhiteSpace(toTable) ||
            string.IsNullOrWhiteSpace(toColumn))
        {
            _console.WriteError("--from-table, --from-column, --to-table, and --to-column are required for this action.");
            return false;
        }

        return true;
    }

    private int WriteResult(ResultBase result)
    {
        _console.WriteJson(result);
        return result.Success ? 0 : -1;
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown datamodel action '{action}'.");
        return -1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--table <TABLE>")]
        public string? TableName { get; init; }

        [CommandOption("--new-name <NAME>")]
        public string? NewName { get; init; }

        [CommandOption("--measure <MEASURE>")]
        public string? MeasureName { get; init; }

        [CommandOption("--dax <FORMULA>")]
        public string? DaxFormula { get; init; }

        [CommandOption("--format-type <FORMAT>")]
        public string? FormatType { get; init; }

        [CommandOption("--description <TEXT>")]
        public string? Description { get; init; }

        [CommandOption("--output <PATH>")]
        public string? OutputPath { get; init; }

        [CommandOption("--from-table <TABLE>")]
        public string? FromTable { get; init; }

        [CommandOption("--from-column <COLUMN>")]
        public string? FromColumn { get; init; }

        [CommandOption("--to-table <TABLE>")]
        public string? ToTable { get; init; }

        [CommandOption("--to-column <COLUMN>")]
        public string? ToColumn { get; init; }

        [CommandOption("--active <BOOL>")]
        public bool? Active { get; init; }

        [CommandOption("--timeout-seconds <SECONDS>")]
        public int? TimeoutSeconds { get; init; }

        [CommandOption("--dax-query <QUERY>")]
        public string? DaxQuery { get; init; }

        [CommandOption("--dmv-query <QUERY>")]
        public string? DmvQuery { get; init; }
    }
}
