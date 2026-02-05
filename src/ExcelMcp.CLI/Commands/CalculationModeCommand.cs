using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Service;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Calculation mode commands - thin wrapper that sends requests to service.
/// Actions: get-mode, set-mode, calculate
/// </summary>
internal sealed class CalculationModeCommand : AsyncCommand<CalculationModeCommand.Settings>
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

        if (!ActionValidator.TryNormalizeAction<CalculationModeAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }

        var command = $"calculation.{action}";
        object? args = action switch
        {
            "get-mode" => null,
            "set-mode" => new { mode = settings.Mode },
            "calculate" => new
            {
                scope = string.IsNullOrWhiteSpace(settings.Scope) ? "workbook" : settings.Scope,
                sheetName = settings.SheetName,
                rangeAddress = settings.RangeAddress
            },
            _ => null
        };

        if (action == "set-mode" && string.IsNullOrWhiteSpace(settings.Mode))
        {
            AnsiConsole.MarkupLine("[red]--mode is required for set-mode (automatic, manual, semi-automatic).[/]");
            return 1;
        }

        if (action == "calculate")
        {
            var scope = string.IsNullOrWhiteSpace(settings.Scope) ? "workbook" : settings.Scope;
            if (scope.Equals("sheet", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(settings.SheetName))
            {
                AnsiConsole.MarkupLine("[red]--sheet is required for calculate --scope sheet.[/]");
                return 1;
            }

            if (scope.Equals("range", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(settings.SheetName) || string.IsNullOrWhiteSpace(settings.RangeAddress))
                {
                    AnsiConsole.MarkupLine("[red]--sheet and --range are required for calculate --scope range.[/]");
                    return 1;
                }
            }
        }

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = command,
            SessionId = settings.SessionId,
            Args = args != null ? JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions) : null
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(!string.IsNullOrEmpty(response.Result)
                ? response.Result
                : JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
            return 0;
        }

        Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
        return 1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        [Description("The action to perform (get-mode, set-mode, calculate)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--mode <MODE>")]
        [Description("Calculation mode for set-mode (automatic, manual, semi-automatic)")]
        public string? Mode { get; init; }

        [CommandOption("--scope <SCOPE>")]
        [Description("Calculation scope for calculate (workbook, sheet, range)")]
        public string? Scope { get; init; }

        [CommandOption("--sheet <NAME>")]
        [Description("Worksheet name for calculate scope sheet/range")]
        public string? SheetName { get; init; }

        [CommandOption("--range <ADDRESS>")]
        [Description("Range address for calculate scope range (e.g., A1:C10)")]
        public string? RangeAddress { get; init; }
    }
}
