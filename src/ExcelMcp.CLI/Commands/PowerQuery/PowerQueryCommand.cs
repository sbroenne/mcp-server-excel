using Spectre.Console.Cli;

using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;

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
            _console.WriteError("Action is required (list, view).");
            return -1;
        }

        var batch = _sessionService.GetBatch(settings.SessionId);

        var exitCode = action switch
        {
            "list" => WriteResult(_powerQueryCommands.List(batch)),
            "view" => ExecuteView(batch, settings),
            _ => ReportUnknown(action)
        };

        return exitCode;
    }

    private int ExecuteView(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.QueryName))
        {
            _console.WriteError("Query name is required for 'view' action (-q|--query).");
            return -1;
        }

        return WriteResult(_powerQueryCommands.View(batch, settings.QueryName));
    }

    private int WriteResult(ResultBase result)
    {
        _console.WriteJson(result);
        return result.Success ? 0 : -1;
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown action '{action}'. Supported actions: list, view.");
        return -1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = "list";

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("-q|--query <NAME>")]
        public string? QueryName { get; init; }
    }
}
