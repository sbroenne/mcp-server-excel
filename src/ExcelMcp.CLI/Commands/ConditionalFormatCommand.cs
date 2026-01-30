using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// ConditionalFormat commands - thin wrapper that sends requests to daemon.
/// Actions: add-rule, clear-rules
/// </summary>
internal sealed class ConditionalFormatCommand : AsyncCommand<ConditionalFormatCommand.Settings>
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

        if (!ActionValidator.TryNormalizeAction<ConditionalFormatAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"conditionalformat.{action}";

        object? args = action switch
        {
            "add-rule" => new { sheetName = settings.SheetName, rangeAddress = settings.Range, ruleType = settings.RuleType, formula = settings.Formula, formatStyle = settings.FormatStyle },
            "clear-rules" => new { sheetName = settings.SheetName, rangeAddress = settings.Range },
            _ => new { sheetName = settings.SheetName, rangeAddress = settings.Range }
        };

        using var client = new DaemonClient();
        var response = await client.SendAsync(new DaemonRequest
        {
            Command = command,
            SessionId = settings.SessionId,
            Args = args != null ? JsonSerializer.Serialize(args, DaemonProtocol.JsonOptions) : null
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(!string.IsNullOrEmpty(response.Result) ? response.Result : JsonSerializer.Serialize(new { success = true }, DaemonProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, DaemonProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--sheet <NAME>")]
        public string? SheetName { get; init; }

        [CommandOption("--range <ADDRESS>")]
        public string? Range { get; init; }

        [CommandOption("--rule-type <TYPE>")]
        public string? RuleType { get; init; }

        [CommandOption("--formula <FORMULA>")]
        public string? Formula { get; init; }

        [CommandOption("--format-style <STYLE>")]
        public string? FormatStyle { get; init; }
    }
}
