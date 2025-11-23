using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.Connection;

internal sealed class ConnectionCommand : Command<ConnectionCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly IConnectionCommands _connectionCommands;
    private readonly ICliConsole _console;

    public ConnectionCommand(ISessionService sessionService, IConnectionCommands connectionCommands, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _connectionCommands = connectionCommands ?? throw new ArgumentNullException(nameof(connectionCommands));
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
            "list" => WriteResult(_connectionCommands.List(batch)),
            "view" => ExecuteView(batch, settings),
            "create" => ExecuteCreate(batch, settings),
            "refresh" => ExecuteRefresh(batch, settings),
            "delete" => ExecuteDelete(batch, settings),
            "load-to" => ExecuteLoadTo(batch, settings),
            "get-properties" => ExecuteGetProperties(batch, settings),
            "set-properties" => ExecuteSetProperties(batch, settings),
            "test" => ExecuteTest(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    private int ExecuteView(IExcelBatch batch, Settings settings)
    {
        if (!TryGetName(settings, out var name))
        {
            return -1;
        }

        return WriteResult(_connectionCommands.View(batch, name));
    }

    private int ExecuteCreate(IExcelBatch batch, Settings settings)
    {
        if (!TryGetName(settings, out var name) || string.IsNullOrWhiteSpace(settings.ConnectionString))
        {
            if (string.IsNullOrWhiteSpace(settings.ConnectionString))
            {
                _console.WriteError("--connection-string is required for create.");
            }
            return -1;
        }

        try
        {
            _connectionCommands.Create(
                batch,
                name,
                settings.ConnectionString!,
                settings.CommandText,
                settings.Description);

            _console.WriteInfo($"Connection '{name}' created successfully.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to create connection '{name}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteRefresh(IExcelBatch batch, Settings settings)
    {
        if (!TryGetName(settings, out var name))
        {
            return -1;
        }

        TimeSpan? timeout = settings.TimeoutSeconds.HasValue ? TimeSpan.FromSeconds(settings.TimeoutSeconds.Value) : null;
        try
        {
            _connectionCommands.Refresh(batch, name, timeout);
            _console.WriteInfo($"Connection '{name}' refreshed successfully.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to refresh connection '{name}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteDelete(IExcelBatch batch, Settings settings)
    {
        if (!TryGetName(settings, out var name))
        {
            return -1;
        }

        try
        {
            _connectionCommands.Delete(batch, name);
            _console.WriteInfo($"Connection '{name}' deleted successfully.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to delete connection '{name}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteLoadTo(IExcelBatch batch, Settings settings)
    {
        if (!TryGetName(settings, out var name) || string.IsNullOrWhiteSpace(settings.SheetName))
        {
            if (string.IsNullOrWhiteSpace(settings.SheetName))
            {
                _console.WriteError("--sheet is required for load-to.");
            }
            return -1;
        }

        try
        {
            _connectionCommands.LoadTo(batch, name, settings.SheetName!);
            _console.WriteInfo($"Connection '{name}' loaded to sheet '{settings.SheetName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to load connection '{name}' to sheet '{settings.SheetName}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteGetProperties(IExcelBatch batch, Settings settings)
    {
        if (!TryGetName(settings, out var name))
        {
            return -1;
        }

        return WriteResult(_connectionCommands.GetProperties(batch, name));
    }

    private int ExecuteSetProperties(IExcelBatch batch, Settings settings)
    {
        if (!TryGetName(settings, out var name))
        {
            return -1;
        }

        if (settings.BackgroundQuery is null &&
            settings.RefreshOnFileOpen is null &&
            settings.SavePassword is null &&
            settings.RefreshPeriodMinutes is null)
        {
            _console.WriteError("Provide at least one property option for set-properties.");
            return -1;
        }

        try
        {
            _connectionCommands.SetProperties(
                batch,
                name,
                connectionString: null,
                commandText: null,
                description: null,
                settings.BackgroundQuery,
                settings.RefreshOnFileOpen,
                settings.SavePassword,
                settings.RefreshPeriodMinutes);

            _console.WriteInfo($"Updated properties for connection '{name}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set properties for connection '{name}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteTest(IExcelBatch batch, Settings settings)
    {
        if (!TryGetName(settings, out var name))
        {
            return -1;
        }

        try
        {
            _connectionCommands.Test(batch, name);
            _console.WriteInfo($"Connection '{name}' passed validation.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Connection '{name}' test failed: {ex.Message}");
            return 1;
        }
    }

    private bool TryGetName(Settings settings, out string name)
    {
        name = settings.ConnectionName?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(name))
        {
            _console.WriteError("--name is required for this action.");
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
        _console.WriteError($"Unknown connection action '{action}'.");
        return -1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--name <NAME>")]
        public string? ConnectionName { get; init; }

        [CommandOption("--connection-string <STRING>")]
        public string? ConnectionString { get; init; }

        [CommandOption("--command-text <COMMAND>")]
        public string? CommandText { get; init; }

        [CommandOption("--description <TEXT>")]
        public string? Description { get; init; }

        [CommandOption("--sheet <SHEET>")]
        public string? SheetName { get; init; }

        [CommandOption("--timeout-seconds <SECONDS>")]
        public int? TimeoutSeconds { get; init; }

        [CommandOption("--background-query <BOOL>")]
        public bool? BackgroundQuery { get; init; }

        [CommandOption("--refresh-on-open <BOOL>")]
        public bool? RefreshOnFileOpen { get; init; }

        [CommandOption("--save-password <BOOL>")]
        public bool? SavePassword { get; init; }

        [CommandOption("--refresh-period-minutes <MINUTES>")]
        public int? RefreshPeriodMinutes { get; init; }
    }
}
