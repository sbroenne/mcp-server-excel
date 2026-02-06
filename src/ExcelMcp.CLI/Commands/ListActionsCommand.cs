using System.ComponentModel;
using System.Text.Json;
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
        var actionsByCommand = new Dictionary<string, IEnumerable<string>>(StringComparer.OrdinalIgnoreCase)
        {
            // Session management (REQUIRED FIRST STEP)
            ["session"] = new[] { "create", "open", "close", "list", "save" },
            // All commands below match 1:1 with MCP tools (excel_ prefix removed)
            ["worksheet"] = ServiceRegistry.Sheet.ValidActions,
            ["worksheetstyle"] = ServiceRegistry.SheetStyle.ValidActions,
            ["range"] = ServiceRegistry.Range.ValidActions,
            ["rangeedit"] = ServiceRegistry.RangeEdit.ValidActions,
            ["rangeformat"] = ServiceRegistry.RangeFormat.ValidActions,
            ["rangelink"] = ServiceRegistry.RangeLink.ValidActions,
            ["table"] = ServiceRegistry.Table.ValidActions,
            ["tablecolumn"] = ServiceRegistry.TableColumn.ValidActions,
            ["powerquery"] = ServiceRegistry.PowerQuery.ValidActions,
            ["pivottable"] = ServiceRegistry.PivotTable.ValidActions,
            ["pivottablefield"] = ServiceRegistry.PivotTableField.ValidActions,
            ["pivottablecalc"] = ServiceRegistry.PivotTableCalc.ValidActions,
            ["chart"] = ServiceRegistry.Chart.ValidActions,
            ["chartconfig"] = ServiceRegistry.ChartConfig.ValidActions,
            ["connection"] = ServiceRegistry.Connection.ValidActions,
            ["calculationmode"] = ServiceRegistry.Calculation.ValidActions,
            ["namedrange"] = ServiceRegistry.NamedRange.ValidActions,
            ["conditionalformat"] = ServiceRegistry.ConditionalFormat.ValidActions,
            ["vba"] = ServiceRegistry.Vba.ValidActions,
            ["datamodel"] = ServiceRegistry.DataModel.ValidActions,
            ["datamodelrel"] = ServiceRegistry.DataModelRel.ValidActions,
            ["slicer"] = ServiceRegistry.Slicer.ValidActions
        };

        if (!string.IsNullOrWhiteSpace(settings.CommandName))
        {
            var key = settings.CommandName.Trim().ToLowerInvariant();
            if (!actionsByCommand.TryGetValue(key, out var actions))
            {
                var error = new { success = false, error = $"Unknown command '{key}'." };
                Console.WriteLine(JsonSerializer.Serialize(error, ServiceProtocol.JsonOptions));
                return 1;
            }

            var result = new
            {
                success = true,
                command = key,
                actions = actions.OrderBy(a => a, StringComparer.OrdinalIgnoreCase).ToArray()
            };
            Console.WriteLine(JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions));
            return 0;
        }

        var all = actionsByCommand.ToDictionary(
            pair => pair.Key,
            pair => pair.Value.OrderBy(a => a, StringComparer.OrdinalIgnoreCase).ToArray(),
            StringComparer.OrdinalIgnoreCase);

        var payload = new
        {
            success = true,
            workflow = "REQUIRED: 1) session open/create <file> → get sessionId, 2) all commands need --session <id>, 3) session close --save to persist",
            example = "session create file.xlsx → returns {sessionId:'abc'} → range set-values --session abc --range A1 --values 'Hello' → session close --save --session abc",
            commands = all
        };
        Console.WriteLine(JsonSerializer.Serialize(payload, ServiceProtocol.JsonOptions));
        return 0;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "[COMMAND]")]
        [Description("Command name to list actions for (omit to list all commands)")]
        public string? CommandName { get; init; }
    }
}


