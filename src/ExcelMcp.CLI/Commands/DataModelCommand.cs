using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Service;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// DataModel commands - thin wrapper that sends requests to service.
/// Actions: list-tables, read-table, list-columns, list-measures, read, create-measure, update-measure, delete-measure,
/// rename-table, delete-table, read-info, refresh, evaluate, execute-dmv
/// </summary>
internal sealed class DataModelCommand : AsyncCommand<DataModelCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            AnsiConsole.MarkupLine("[red]Session ID is required. Use --session <id>[/]");
            return 1;
        }

        if (string.IsNullOrWhiteSpace(settings.Action))
        {
            AnsiConsole.MarkupLine("[red]Action is required.[/]");
            return 1;
        }

        if (!ActionValidator.TryNormalizeAction<DataModelAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"datamodel.{action}";

        // Note: property names must match daemon's Args classes (e.g., DataModelCreateMeasureArgs)
        var expression = ResolveFileOrValue(settings.Expression, settings.ExpressionFile);
        var daxQuery = ResolveFileOrValue(settings.DaxQuery, settings.DaxQueryFile);
        var dmvQuery = ResolveFileOrValue(settings.DmvQuery, settings.DmvQueryFile);
        object? args = action switch
        {
            "list-tables" => null,
            "read-table" => new { tableName = settings.TableName, maxRows = settings.MaxRows },
            "list-columns" => new { tableName = settings.TableName },
            "list-measures" => new { tableName = settings.TableName },
            "read" => new { measureName = settings.MeasureName, tableName = settings.TableName },
            "create-measure" => new { tableName = settings.TableName, measureName = settings.MeasureName, daxFormula = expression, formatType = settings.FormatString },
            "update-measure" => new { measureName = settings.MeasureName, tableName = settings.TableName, daxFormula = expression, formatType = settings.FormatString },
            "delete-measure" => new { measureName = settings.MeasureName, tableName = settings.TableName },
            "rename-table" => new { oldName = settings.TableName, newName = settings.NewName },
            "delete-table" => new { tableName = settings.TableName },
            "read-info" => null,
            "refresh" => null,
            "evaluate" => new { daxQuery },
            "execute-dmv" => new { dmvQuery },
            _ => new { tableName = settings.TableName }
        };

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = command,
            SessionId = settings.SessionId,
            Args = args != null ? JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions) : null
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(!string.IsNullOrEmpty(response.Result) ? response.Result : JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    /// <summary>
    /// Returns file contents if filePath is provided, otherwise returns the direct value.
    /// </summary>
    private static string? ResolveFileOrValue(string? directValue, string? filePath)
    {
        if (!string.IsNullOrWhiteSpace(filePath))
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found: {filePath}");
            }
            return File.ReadAllText(filePath);
        }
        return directValue;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        [Description("The action to perform (e.g., list-tables, create-measure, evaluate)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--table <NAME>")]
        [Description("Data Model table name")]
        public string? TableName { get; init; }

        [CommandOption("--measure <NAME>")]
        [Description("Measure name")]
        public string? MeasureName { get; init; }

        [CommandOption("--new-name <NAME>")]
        [Description("New name for rename operation")]
        public string? NewName { get; init; }

        [CommandOption("--expression <DAX>")]
        [Description("DAX formula for measure")]
        public string? Expression { get; init; }

        [CommandOption("--expression-file <PATH>")]
        [Description("Read DAX formula from file instead of command line")]
        public string? ExpressionFile { get; init; }

        [CommandOption("--format-string <FORMAT>")]
        [Description("Number format string for measure")]
        public string? FormatString { get; init; }

        [CommandOption("--dax-query <QUERY>")]
        [Description("DAX query to evaluate")]
        public string? DaxQuery { get; init; }

        [CommandOption("--dax-query-file <PATH>")]
        [Description("Read DAX query from file instead of command line")]
        public string? DaxQueryFile { get; init; }

        [CommandOption("--dmv-query <QUERY>")]
        [Description("DMV query to execute")]
        public string? DmvQuery { get; init; }

        [CommandOption("--dmv-query-file <PATH>")]
        [Description("Read DMV query from file instead of command line")]
        public string? DmvQueryFile { get; init; }

        [CommandOption("--max-rows <COUNT>")]
        [Description("Maximum rows to return")]
        public int? MaxRows { get; init; }
    }
}
