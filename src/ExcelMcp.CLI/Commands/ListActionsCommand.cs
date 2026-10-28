using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Generated;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Lists available actions for CLI commands.
/// </summary>
internal sealed class ListActionsCommand : Command<ListActionsCommand.Settings>
{
    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        // Session actions are hand-maintained (bootstrap operations, not generated)
        // All other commands use the generated ValidActionsByCommand mapping
        var actionsByCommand = new Dictionary<string, IEnumerable<string>>(
            _CliCategoryMetadata.ValidActionsByCommand.Select(kv =>
                new KeyValuePair<string, IEnumerable<string>>(kv.Key, kv.Value)),
            StringComparer.OrdinalIgnoreCase)
        {
            ["session"] = new[] { "create", "open", "close", "list", "save", "discover", "attach" }
        };

        if (!string.IsNullOrWhiteSpace(settings.CommandName))
        {
            var key = settings.CommandName.Trim().ToLowerInvariant();
            if (!actionsByCommand.TryGetValue(key, out var actions))
            {
                var error = new { success = false, errorMessage = $"Unknown command '{key}'." };
                return CliResultOutput.WriteJson(JsonSerializer.Serialize(error, ServiceProtocol.JsonOptions), "list-actions");
            }

            var result = new
            {
                success = true,
                command = key,
                actions = actions.OrderBy(a => a, StringComparer.OrdinalIgnoreCase).ToArray()
            };
            return CliResultOutput.WriteJson(JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions), "list-actions");
        }

        var all = actionsByCommand.ToDictionary(
            pair => pair.Key,
            pair => pair.Value.OrderBy(a => a, StringComparer.OrdinalIgnoreCase).ToArray(),
            StringComparer.OrdinalIgnoreCase);

        var payload = new
        {
            success = true,
            workflow = "1) session open/create <file>, or session discover then session attach --workbook-id <id>, to get sessionId; 2) workbook commands use --session <id>; 3) session save saves explicitly; session close closes the selected workbook (add --save to save changes). Service shutdown leaves newly attached workbooks open.",
            example = "session create file.xlsx → returns {sessionId:'abc'} → range set-values --session abc --range-address A1 --values 'Hello' → session close --save --session abc",
            commands = all
        };
        return CliResultOutput.WriteJson(JsonSerializer.Serialize(payload, ServiceProtocol.JsonOptions), "list-actions");
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "[COMMAND]")]
        [Description("Command name to list actions for (omit to list all commands)")]
        public string? CommandName { get; init; }
    }
}
