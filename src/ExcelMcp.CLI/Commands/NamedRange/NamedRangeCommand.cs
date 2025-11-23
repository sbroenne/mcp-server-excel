using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.NamedRange;

internal sealed class NamedRangeCommand : Command<NamedRangeCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly INamedRangeCommands _namedRangeCommands;
    private readonly ICliConsole _console;

    public NamedRangeCommand(ISessionService sessionService, INamedRangeCommands namedRangeCommands, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _namedRangeCommands = namedRangeCommands ?? throw new ArgumentNullException(nameof(namedRangeCommands));
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
            "list" => ExecuteList(batch),
            "get" => ExecuteGet(batch, settings),
            "set" => ExecuteSet(batch, settings),
            "create" => ExecuteCreate(batch, settings),
            "update" => ExecuteUpdate(batch, settings),
            "delete" => ExecuteDelete(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    private int ExecuteList(IExcelBatch batch)
    {
        try
        {
            var namedRanges = _namedRangeCommands.List(batch);
            _console.WriteJson(namedRanges);
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to list named ranges: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteGet(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.Name))
        {
            _console.WriteError("--name is required for get.");
            return -1;
        }

        try
        {
            var value = _namedRangeCommands.Read(batch, settings.Name);
            _console.WriteJson(value);
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to get named range '{settings.Name}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSet(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.Name) || settings.Value is null)
        {
            _console.WriteError("--name and --value are required for set.");
            return -1;
        }

        try
        {
            _namedRangeCommands.Write(batch, settings.Name, settings.Value);
            _console.WriteInfo($"Named range '{settings.Name}' value updated successfully.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set named range '{settings.Name}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteCreate(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.Name) || string.IsNullOrWhiteSpace(settings.Reference))
        {
            _console.WriteError("--name and --reference are required for create.");
            return -1;
        }

        try
        {
            _namedRangeCommands.Create(batch, settings.Name, settings.Reference);
            _console.WriteInfo($"Named range '{settings.Name}' created successfully.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to create named range '{settings.Name}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteUpdate(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.Name) || string.IsNullOrWhiteSpace(settings.Reference))
        {
            _console.WriteError("--name and --reference are required for update.");
            return -1;
        }

        try
        {
            _namedRangeCommands.Update(batch, settings.Name, settings.Reference);
            _console.WriteInfo($"Named range '{settings.Name}' updated successfully.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to update named range '{settings.Name}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteDelete(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.Name))
        {
            _console.WriteError("--name is required for delete.");
            return -1;
        }

        try
        {
            _namedRangeCommands.Delete(batch, settings.Name);
            _console.WriteInfo($"Named range '{settings.Name}' deleted successfully.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to delete named range '{settings.Name}': {ex.Message}");
            return 1;
        }
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown named range action '{action}'.");
        return -1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--name <NAME>")]
        public string? Name { get; init; }

        [CommandOption("--value <VALUE>")]
        public string? Value { get; init; }

        [CommandOption("--reference <REFERENCE>")]
        public string? Reference { get; init; }

        [CommandOption("--definitions-json <JSON>")]
        public string? DefinitionsJson { get; init; }

        [CommandOption("--definitions-file <PATH>")]
        public string? DefinitionsFile { get; init; }
    }
}
