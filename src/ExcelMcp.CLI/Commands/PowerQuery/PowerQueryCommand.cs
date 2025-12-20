using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.PowerQuery;

internal sealed class PowerQueryCommand : Command<PowerQueryCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly IPowerQueryCommands _powerQueryCommands;
    private readonly ICliConsole _console;

    public PowerQueryCommand(
        ISessionService sessionService,
        IPowerQueryCommands powerQueryCommands,
        ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _powerQueryCommands = powerQueryCommands ?? throw new ArgumentNullException(nameof(powerQueryCommands));
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
            _console.WriteError("Action is required (list, view, rename).");
            return -1;
        }

        if (!string.IsNullOrWhiteSpace(settings.TargetCellAddress) && action is not ("create" or "load-to"))
        {
            _console.WriteError("--target-cell is only supported for 'create' and 'load-to' actions.");
            return -1;
        }

        var batch = _sessionService.GetBatch(settings.SessionId);

        var exitCode = action switch
        {
            "list" => WriteResult(_powerQueryCommands.List(batch)),
            "view" => ExecuteView(batch, settings),
            "create" => ExecuteCreate(batch, settings),
            "update" => ExecuteUpdate(batch, settings),
            "delete" => ExecuteDelete(batch, settings),
            "rename" => ExecuteRename(batch, settings),
            "refresh" => ExecuteRefresh(batch, settings),
            "get-load-config" => ExecuteGetLoadConfig(batch, settings),
            "refresh-all" => ExecuteRefreshAll(batch),
            "load-to" => ExecuteLoadTo(batch, settings),
            _ => ReportUnknown(action)
        };

        return exitCode;
    }

    private int ExecuteView(IExcelBatch batch, Settings settings)
    {
        if (!TryGetQueryName(settings, out var queryName))
        {
            return -1;
        }

        return WriteResult(_powerQueryCommands.View(batch, queryName));
    }

    private int ExecuteCreate(IExcelBatch batch, Settings settings)
    {
        if (!TryGetQueryName(settings, out var queryName) ||
            !TryReadMCode(settings.MCodeFile, out var mCode) ||
            !TryParseLoadMode(settings.LoadDestination, out var loadMode))
        {
            return -1;
        }

        if (!RequiresWorksheet(loadMode) && !string.IsNullOrWhiteSpace(settings.TargetCellAddress))
        {
            _console.WriteError("--target-cell can only be used when load destination is 'worksheet' or 'both'.");
            return -1;
        }

        try
        {
            _powerQueryCommands.Create(
                batch,
                queryName,
                mCode,
                loadMode,
                settings.TargetSheet,
                settings.TargetCellAddress);
            _console.WriteJson(new { success = true, message = $"Query '{queryName}' created successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to create query: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteUpdate(IExcelBatch batch, Settings settings)
    {
        if (!TryGetQueryName(settings, out var queryName) ||
            !TryReadMCode(settings.MCodeFile, out var mCode))
        {
            return -1;
        }

        try
        {
            _powerQueryCommands.Update(batch, queryName, mCode);
            _console.WriteJson(new { success = true, message = $"Query '{queryName}' updated successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to update query: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteDelete(IExcelBatch batch, Settings settings)
    {
        if (!TryGetQueryName(settings, out var queryName))
        {
            return -1;
        }

        try
        {
            _powerQueryCommands.Delete(batch, queryName);
            _console.WriteJson(new { success = true, message = $"Query '{queryName}' deleted successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to delete query: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteRename(IExcelBatch batch, Settings settings)
    {
        if (!TryGetQueryName(settings, out var queryName))
        {
            return -1;
        }

        var newName = settings.NewName?.Trim();
        if (string.IsNullOrWhiteSpace(newName))
        {
            _console.WriteError("New name is required (--new-name).");
            return -1;
        }

        return WriteResult(_powerQueryCommands.Rename(batch, queryName, newName));
    }

    private int ExecuteRefresh(IExcelBatch batch, Settings settings)
    {
        if (!TryGetQueryName(settings, out var queryName))
        {
            return -1;
        }

        var timeout = TimeSpan.FromSeconds(settings.RefreshTimeoutSeconds ?? 60);
        return WriteResult(_powerQueryCommands.Refresh(batch, queryName, timeout));
    }

    private int ExecuteGetLoadConfig(IExcelBatch batch, Settings settings)
    {
        if (!TryGetQueryName(settings, out var queryName))
        {
            return -1;
        }

        return WriteResult(_powerQueryCommands.GetLoadConfig(batch, queryName));
    }

    private int ExecuteRefreshAll(IExcelBatch batch)
    {
        try
        {
            _powerQueryCommands.RefreshAll(batch);
            _console.WriteJson(new { success = true, message = "All queries refreshed successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to refresh queries: {ex.Message}");
            return -1;
        }
    }

    private int ExecuteLoadTo(IExcelBatch batch, Settings settings)
    {
        if (!TryGetQueryName(settings, out var queryName) ||
            !TryParseLoadMode(settings.LoadDestination, out var loadMode))
        {
            return -1;
        }

        if (!RequiresWorksheet(loadMode) && !string.IsNullOrWhiteSpace(settings.TargetCellAddress))
        {
            _console.WriteError("--target-cell can only be used when load destination is 'worksheet' or 'both'.");
            return -1;
        }

        string? targetSheet = settings.TargetSheet;
        if (RequiresWorksheet(loadMode) && string.IsNullOrWhiteSpace(targetSheet))
        {
            targetSheet = queryName;
        }

        try
        {
            _powerQueryCommands.LoadTo(
                batch,
                queryName,
                loadMode,
                targetSheet,
                settings.TargetCellAddress);
            _console.WriteJson(new { success = true, message = $"Query '{queryName}' load configuration applied successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to apply load configuration: {ex.Message}");
            return -1;
        }
    }

    private int WriteResult(ResultBase result)
    {
        _console.WriteJson(result);
        return result.Success ? 0 : -1;
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown action '{action}'. Supported actions: list, view, create, update, delete, rename, refresh, get-load-config, refresh-all, load-to.");
        return -1;
    }

    private bool TryGetQueryName(Settings settings, out string queryName)
    {
        queryName = settings.QueryName?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(queryName))
        {
            _console.WriteError("Query name is required (-q|--query).");
            return false;
        }

        return true;
    }

    private bool TryReadMCode(string? path, out string mCode)
    {
        mCode = string.Empty;
        if (string.IsNullOrWhiteSpace(path))
        {
            _console.WriteError("--m-file is required for this action.");
            return false;
        }

        if (!System.IO.File.Exists(path))
        {
            _console.WriteError($"M code file '{path}' was not found.");
            return false;
        }

        mCode = System.IO.File.ReadAllText(path);
        return true;
    }

    private bool TryParseLoadMode(string? loadDestination, out PowerQueryLoadMode loadMode)
    {
        var value = loadDestination?.Trim();
        if (string.IsNullOrEmpty(value))
        {
            loadMode = PowerQueryLoadMode.LoadToTable;
            return true;
        }

        switch (value.ToLowerInvariant())
        {
            case "worksheet":
                loadMode = PowerQueryLoadMode.LoadToTable;
                return true;
            case "data-model":
                loadMode = PowerQueryLoadMode.LoadToDataModel;
                return true;
            case "both":
                loadMode = PowerQueryLoadMode.LoadToBoth;
                return true;
            case "connection-only":
                loadMode = PowerQueryLoadMode.ConnectionOnly;
                return true;
            default:
                _console.WriteError("--load-destination must be one of: worksheet, data-model, both, connection-only.");
                loadMode = PowerQueryLoadMode.LoadToTable;
                return false;
        }
    }

    private static bool RequiresWorksheet(PowerQueryLoadMode loadMode)
    {
        return loadMode == PowerQueryLoadMode.LoadToTable || loadMode == PowerQueryLoadMode.LoadToBoth;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = "list";

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("-q|--query <NAME>")]
        public string? QueryName { get; init; }

        [CommandOption("--new-name <NAME>")]
        public string? NewName { get; init; }

        [CommandOption("--m-file <PATH>")]
        public string? MCodeFile { get; init; }

        [CommandOption("--load-destination <MODE>")]
        public string? LoadDestination { get; init; }

        [CommandOption("--target-sheet <NAME>")]
        public string? TargetSheet { get; init; }

        [CommandOption("--target-cell <ADDRESS>")]
        public string? TargetCellAddress { get; init; }

        [CommandOption("--refresh-timeout <SECONDS>")]
        public int? RefreshTimeoutSeconds { get; init; }
    }
}
