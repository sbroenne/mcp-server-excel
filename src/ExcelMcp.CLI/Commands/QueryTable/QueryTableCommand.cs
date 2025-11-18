using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.QueryTable;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.QueryTable;

internal sealed class QueryTableCommand : Command<QueryTableCommand.Settings>
{
    private static readonly JsonSerializerOptions OptionsSerializer = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly ISessionService _sessionService;
    private readonly IQueryTableCommands _queryTableCommands;
    private readonly ICliConsole _console;

    public QueryTableCommand(ISessionService sessionService, IQueryTableCommands queryTableCommands, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _queryTableCommands = queryTableCommands ?? throw new ArgumentNullException(nameof(queryTableCommands));
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
            "list" => WriteResult(_queryTableCommands.List(batch)),
            "get" => ExecuteGet(batch, settings),
            "create-from-connection" => ExecuteCreateFromConnection(batch, settings),
            "create-from-query" => ExecuteCreateFromQuery(batch, settings),
            "refresh" => ExecuteRefresh(batch, settings),
            "refresh-all" => ExecuteRefreshAll(batch, settings),
            "update" => ExecuteUpdate(batch, settings),
            "delete" => ExecuteDelete(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    private int ExecuteGet(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.QueryTableName))
        {
            _console.WriteError("--name is required for get.");
            return -1;
        }

        return WriteResult(_queryTableCommands.Read(batch, settings.QueryTableName));
    }

    private int ExecuteCreateFromConnection(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName) ||
            string.IsNullOrWhiteSpace(settings.QueryTableName) ||
            string.IsNullOrWhiteSpace(settings.ConnectionName))
        {
            _console.WriteError("--sheet, --name, and --connection are required for create-from-connection.");
            return -1;
        }

        var options = LoadCreateOptions(settings);
        if (options == null && (!string.IsNullOrWhiteSpace(settings.CreateOptionsJson) || !string.IsNullOrWhiteSpace(settings.CreateOptionsFile)))
        {
            return -1;
        }

        return WriteResult(_queryTableCommands.CreateFromConnection(
            batch,
            settings.SheetName,
            settings.QueryTableName,
            settings.ConnectionName,
            settings.Range ?? "A1",
            options));
    }

    private int ExecuteCreateFromQuery(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName) ||
            string.IsNullOrWhiteSpace(settings.QueryTableName) ||
            string.IsNullOrWhiteSpace(settings.QueryName))
        {
            _console.WriteError("--sheet, --name, and --query are required for create-from-query.");
            return -1;
        }

        var options = LoadCreateOptions(settings);
        if (options == null && (!string.IsNullOrWhiteSpace(settings.CreateOptionsJson) || !string.IsNullOrWhiteSpace(settings.CreateOptionsFile)))
        {
            return -1;
        }

        return WriteResult(_queryTableCommands.CreateFromQuery(
            batch,
            settings.SheetName,
            settings.QueryTableName,
            settings.QueryName,
            settings.Range ?? "A1",
            options));
    }

    private int ExecuteRefresh(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.QueryTableName))
        {
            _console.WriteError("--name is required for refresh.");
            return -1;
        }

        return WriteResult(_queryTableCommands.Refresh(batch, settings.QueryTableName, GetTimeout(settings)));
    }

    private int ExecuteRefreshAll(IExcelBatch batch, Settings settings)
    {
        return WriteResult(_queryTableCommands.RefreshAll(batch, GetTimeout(settings)));
    }

    private int ExecuteUpdate(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.QueryTableName))
        {
            _console.WriteError("--name is required for update.");
            return -1;
        }

        var options = LoadUpdateOptions(settings);
        if (options == null)
        {
            _console.WriteError("Provide update settings using --update-options-json or --update-options-file.");
            return -1;
        }

        return WriteResult(_queryTableCommands.UpdateProperties(batch, settings.QueryTableName, options));
    }

    private int ExecuteDelete(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.QueryTableName))
        {
            _console.WriteError("--name is required for delete.");
            return -1;
        }

        return WriteResult(_queryTableCommands.Delete(batch, settings.QueryTableName));
    }

    private QueryTableCreateOptions? LoadCreateOptions(Settings settings)
    {
        if (!string.IsNullOrWhiteSpace(settings.CreateOptionsJson))
        {
            var parsed = ParseCreateOptions(settings.CreateOptionsJson!);
            if (parsed == null)
            {
                _console.WriteError("Unable to parse --create-options-json content.");
            }

            return parsed;
        }

        if (!string.IsNullOrWhiteSpace(settings.CreateOptionsFile))
        {
            if (!System.IO.File.Exists(settings.CreateOptionsFile))
            {
                _console.WriteError($"Create options file '{settings.CreateOptionsFile}' was not found.");
                return null;
            }

            var contents = System.IO.File.ReadAllText(settings.CreateOptionsFile);
            var parsed = ParseCreateOptions(contents);
            if (parsed == null)
            {
                _console.WriteError($"Unable to parse JSON from '{settings.CreateOptionsFile}'.");
            }

            return parsed;
        }

        return null;
    }

    private QueryTableUpdateOptions? LoadUpdateOptions(Settings settings)
    {
        if (!string.IsNullOrWhiteSpace(settings.UpdateOptionsJson))
        {
            return ParseUpdateOptions(settings.UpdateOptionsJson!);
        }

        if (!string.IsNullOrWhiteSpace(settings.UpdateOptionsFile))
        {
            if (!System.IO.File.Exists(settings.UpdateOptionsFile))
            {
                _console.WriteError($"Update options file '{settings.UpdateOptionsFile}' was not found.");
                return null;
            }

            var contents = System.IO.File.ReadAllText(settings.UpdateOptionsFile);
            return ParseUpdateOptions(contents);
        }

        return null;
    }

    private static QueryTableCreateOptions? ParseCreateOptions(string json)
    {
        try
        {
            return JsonSerializer.Deserialize<QueryTableCreateOptions>(json, OptionsSerializer);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private static QueryTableUpdateOptions? ParseUpdateOptions(string json)
    {
        try
        {
            return JsonSerializer.Deserialize<QueryTableUpdateOptions>(json, OptionsSerializer);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private static TimeSpan? GetTimeout(Settings settings)
    {
        return settings.TimeoutSeconds.HasValue
            ? TimeSpan.FromSeconds(settings.TimeoutSeconds.Value)
            : null;
    }

    private int WriteResult(ResultBase result)
    {
        _console.WriteJson(result);
        return result.Success ? 0 : -1;
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown querytable action '{action}'.");
        return -1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--sheet <SHEET>")]
        public string? SheetName { get; init; }

        [CommandOption("--name <NAME>")]
        public string? QueryTableName { get; init; }

        [CommandOption("--connection <NAME>")]
        public string? ConnectionName { get; init; }

        [CommandOption("--query <NAME>")]
        public string? QueryName { get; init; }

        [CommandOption("--range <RANGE>")]
        public string? Range { get; init; }

        [CommandOption("--create-options-json <JSON>")]
        public string? CreateOptionsJson { get; init; }

        [CommandOption("--create-options-file <PATH>")]
        public string? CreateOptionsFile { get; init; }

        [CommandOption("--update-options-json <JSON>")]
        public string? UpdateOptionsJson { get; init; }

        [CommandOption("--update-options-file <PATH>")]
        public string? UpdateOptionsFile { get; init; }

        [CommandOption("--timeout-seconds <SECONDS>")]
        public int? TimeoutSeconds { get; init; }
    }
}
