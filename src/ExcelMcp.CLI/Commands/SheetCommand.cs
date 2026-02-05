using System.ComponentModel;
using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Generated;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Sheet commands - thin wrapper that sends requests to service.
/// </summary>
internal sealed class SheetCommand : AsyncCommand<SheetCommand.Settings>
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

        // Validate and normalize action
        var validActions = ServiceRegistry.Sheet.ValidActions
            .Concat(ServiceRegistry.SheetStyle.ValidActions)
            .ToArray();

        var action = settings.Action.Trim().ToLowerInvariant();
        if (!validActions.Contains(action, StringComparer.OrdinalIgnoreCase))
        {
            var validList = string.Join(", ", validActions);
            AnsiConsole.MarkupLine($"[red]Invalid action '{action}'. Valid actions: {validList}[/]");
            return 1;
        }
        var command = $"sheet.{action}";

        // For set-tab-color, validate hex color first (if provided)
        if (action == "set-tab-color" && !string.IsNullOrWhiteSpace(settings.Color))
        {
            if (!TryParseHexColor(settings.Color, out _, out _, out _))
            {
                AnsiConsole.MarkupLine($"[red]Error: Invalid hex color '{settings.Color}'. Expected format: #RRGGBB (e.g., #FFD966) or #RGB (e.g., #FD9)[/]");
                return 1;
            }
        }

        // Build args based on action
        object? args = action switch
        {
            // WorksheetAction
            "list" => null,
            "create" => new { sheetName = settings.SheetName },
            "rename" => new { sheetName = settings.SheetName, newName = settings.NewName },
            "delete" => new { sheetName = settings.SheetName },
            "copy" => new { sourceSheet = settings.SourceSheet, targetSheet = settings.TargetSheet },
            "move" => new { sheetName = settings.SheetName, beforeSheet = settings.BeforeSheet, afterSheet = settings.AfterSheet },
            "copy-to-file" => new { sourceFile = settings.SourceFile, sourceSheet = settings.SourceSheet, targetFile = settings.TargetFile, targetSheetName = settings.TargetSheet, beforeSheet = settings.BeforeSheet, afterSheet = settings.AfterSheet },
            "move-to-file" => new { sourceFile = settings.SourceFile, sourceSheet = settings.SourceSheet, targetFile = settings.TargetFile, beforeSheet = settings.BeforeSheet, afterSheet = settings.AfterSheet },

            // WorksheetStyleAction
            "set-tab-color" => BuildSetTabColorArgs(settings),
            "get-tab-color" => new { sheetName = settings.SheetName },
            "clear-tab-color" => new { sheetName = settings.SheetName },
            "hide" => new { sheetName = settings.SheetName },
            "very-hide" => new { sheetName = settings.SheetName },
            "show" => new { sheetName = settings.SheetName },
            "get-visibility" => new { sheetName = settings.SheetName },
            "set-visibility" => new { sheetName = settings.SheetName, visibility = settings.Visibility },

            _ => null
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
            if (!string.IsNullOrEmpty(response.Result))
            {
                Console.WriteLine(response.Result);
            }
            else
            {
                Console.WriteLine(JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
            }
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    /// <summary>
    /// Tries to parse a hex color string (#RRGGBB or #RGB) into RGB components.
    /// </summary>
    private static bool TryParseHexColor(string color, out int red, out int green, out int blue)
    {
        red = green = blue = 0;
        var hex = color.TrimStart('#');

        // Support 3-char shorthand (#RGB → #RRGGBB)
        if (hex.Length == 3)
        {
            hex = $"{hex[0]}{hex[0]}{hex[1]}{hex[1]}{hex[2]}{hex[2]}";
        }

        if (hex.Length != 6)
        {
            return false;
        }

        return int.TryParse(hex[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out red) &&
               int.TryParse(hex[2..4], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out green) &&
               int.TryParse(hex[4..6], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out blue);
    }

    /// <summary>
    /// Builds the set-tab-color args.
    /// If --color is provided, it takes precedence over --red/--green/--blue.
    /// Assumes hex color validation was already done.
    /// </summary>
    private static object BuildSetTabColorArgs(Settings settings)
    {
        int? red = settings.Red;
        int? green = settings.Green;
        int? blue = settings.Blue;

        // If --color hex is provided, use it (takes precedence over individual RGB)
        if (!string.IsNullOrWhiteSpace(settings.Color) &&
            TryParseHexColor(settings.Color, out var r, out var g, out var b))
        {
            red = r;
            green = g;
            blue = b;
        }

        return new { sheetName = settings.SheetName, red, green, blue };
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        [Description("The action to perform (e.g., list, create, rename, delete, copy, move)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--sheet <NAME>")]
        [Description("Worksheet name")]
        public string? SheetName { get; init; }

        [CommandOption("--new-name <NAME>")]
        [Description("New name for rename operation")]
        public string? NewName { get; init; }

        [CommandOption("--source-sheet <NAME>")]
        [Description("Source worksheet for copy operation")]
        public string? SourceSheet { get; init; }

        [CommandOption("--target-sheet <NAME>")]
        [Description("Target worksheet name for copy operation")]
        public string? TargetSheet { get; init; }

        [CommandOption("--source-file <PATH>")]
        [Description("Source file path for copy-to-file/move-to-file")]
        public string? SourceFile { get; init; }

        [CommandOption("--target-file <PATH>")]
        [Description("Target file path for copy-to-file/move-to-file")]
        public string? TargetFile { get; init; }

        [CommandOption("--before-sheet <NAME>")]
        [Description("Place sheet before this sheet (move)")]
        public string? BeforeSheet { get; init; }

        [CommandOption("--after-sheet <NAME>")]
        [Description("Place sheet after this sheet (move)")]
        public string? AfterSheet { get; init; }

        [CommandOption("--red <VALUE>")]
        [Description("Red component (0-255) for tab color")]
        public int? Red { get; init; }

        [CommandOption("--green <VALUE>")]
        [Description("Green component (0-255) for tab color")]
        public int? Green { get; init; }

        [CommandOption("--blue <VALUE>")]
        [Description("Blue component (0-255) for tab color")]
        public int? Blue { get; init; }

        [CommandOption("--color <HEX>")]
        [Description("Hex color for tab color (e.g., #FFD966). Overrides --red/--green/--blue if provided.")]
        public string? Color { get; init; }

        [CommandOption("--visibility <VALUE>")]
        [Description("Visibility state: Visible, Hidden, VeryHidden")]
        public string? Visibility { get; init; }
    }
}


