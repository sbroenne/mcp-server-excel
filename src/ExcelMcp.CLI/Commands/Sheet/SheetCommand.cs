using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.Sheet;

internal sealed class SheetCommand : Command<SheetCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly ISheetCommands _sheetCommands;
    private readonly ICliConsole _console;

    public SheetCommand(ISessionService sessionService, ISheetCommands sheetCommands, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _sheetCommands = sheetCommands ?? throw new ArgumentNullException(nameof(sheetCommands));
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
            "list" => WriteResult(_sheetCommands.List(batch)),
            "create" => ExecuteCreate(batch, settings),
            "rename" => ExecuteRename(batch, settings),
            "copy" => ExecuteCopy(batch, settings),
            "delete" => ExecuteDelete(batch, settings),
            "set-tab-color" => ExecuteSetTabColor(batch, settings),
            "get-tab-color" => ExecuteGetTabColor(batch, settings),
            "clear-tab-color" => ExecuteClearTabColor(batch, settings),
            "set-visibility" => ExecuteSetVisibility(batch, settings),
            "show" => ExecuteShow(batch, settings),
            "hide" => ExecuteHide(batch, settings),
            "very-hide" => ExecuteVeryHide(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    private int ExecuteCreate(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("--sheet is required for create.");
            return -1;
        }

        try
        {
            _sheetCommands.Create(batch, settings.SheetName);
            _console.WriteJson(new { success = true, message = $"Sheet '{settings.SheetName}' created successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteJson(new { success = false, message = $"Failed to create sheet '{settings.SheetName}': {ex.Message}" });
            return 1;
        }
    }

    private int ExecuteRename(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName) || string.IsNullOrWhiteSpace(settings.NewSheetName))
        {
            _console.WriteError("--sheet and --new-name are required for rename.");
            return -1;
        }

        try
        {
            _sheetCommands.Rename(batch, settings.SheetName, settings.NewSheetName);
            _console.WriteJson(new { success = true, message = $"Sheet '{settings.SheetName}' renamed to '{settings.NewSheetName}' successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteJson(new { success = false, message = $"Failed to rename sheet '{settings.SheetName}': {ex.Message}" });
            return 1;
        }
    }

    private int ExecuteCopy(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SourceSheet) || string.IsNullOrWhiteSpace(settings.TargetSheet))
        {
            _console.WriteError("--source-sheet and --target-sheet are required for copy.");
            return -1;
        }

        try
        {
            _sheetCommands.Copy(batch, settings.SourceSheet, settings.TargetSheet);
            _console.WriteJson(new { success = true, message = $"Sheet '{settings.SourceSheet}' copied to '{settings.TargetSheet}' successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteJson(new { success = false, message = $"Failed to copy sheet '{settings.SourceSheet}': {ex.Message}" });
            return 1;
        }
    }

    private int ExecuteDelete(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("--sheet is required for delete.");
            return -1;
        }

        try
        {
            _sheetCommands.Delete(batch, settings.SheetName);
            _console.WriteJson(new { success = true, message = $"Sheet '{settings.SheetName}' deleted successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteJson(new { success = false, message = $"Failed to delete sheet '{settings.SheetName}': {ex.Message}" });
            return 1;
        }
    }

    private int ExecuteSetTabColor(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("--sheet is required for set-tab-color.");
            return -1;
        }

        if (!settings.Red.HasValue || !settings.Green.HasValue || !settings.Blue.HasValue)
        {
            _console.WriteError("--red, --green, and --blue are required for set-tab-color.");
            return -1;
        }

        try
        {
            _sheetCommands.SetTabColor(batch, settings.SheetName, settings.Red.Value, settings.Green.Value, settings.Blue.Value);
            _console.WriteJson(new { success = true, message = $"Tab color for sheet '{settings.SheetName}' set successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteJson(new { success = false, message = $"Failed to set tab color for sheet '{settings.SheetName}': {ex.Message}" });
            return 1;
        }
    }

    private int ExecuteGetTabColor(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("--sheet is required for get-tab-color.");
            return -1;
        }

        return WriteResult(_sheetCommands.GetTabColor(batch, settings.SheetName));
    }

    private int ExecuteClearTabColor(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("--sheet is required for clear-tab-color.");
            return -1;
        }

        try
        {
            _sheetCommands.ClearTabColor(batch, settings.SheetName);
            _console.WriteJson(new { success = true, message = $"Tab color for sheet '{settings.SheetName}' cleared successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteJson(new { success = false, message = $"Failed to clear tab color for sheet '{settings.SheetName}': {ex.Message}" });
            return 1;
        }
    }

    private int ExecuteSetVisibility(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("--sheet is required for set-visibility.");
            return -1;
        }

        if (!TryParseVisibility(settings.Visibility, out var visibility))
        {
            _console.WriteError("--visibility must be visible, hidden, or very-hidden.");
            return -1;
        }

        try
        {
            _sheetCommands.SetVisibility(batch, settings.SheetName, visibility);
            _console.WriteJson(new { success = true, message = $"Sheet '{settings.SheetName}' visibility set to {visibility} successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteJson(new { success = false, message = $"Failed to set visibility for sheet '{settings.SheetName}': {ex.Message}" });
            return 1;
        }
    }

    private int ExecuteShow(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("--sheet is required for show.");
            return -1;
        }

        try
        {
            _sheetCommands.Show(batch, settings.SheetName);
            _console.WriteJson(new { success = true, message = $"Sheet '{settings.SheetName}' shown successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteJson(new { success = false, message = $"Failed to show sheet '{settings.SheetName}': {ex.Message}" });
            return 1;
        }
    }

    private int ExecuteHide(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("--sheet is required for hide.");
            return -1;
        }

        try
        {
            _sheetCommands.Hide(batch, settings.SheetName);
            _console.WriteJson(new { success = true, message = $"Sheet '{settings.SheetName}' hidden successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteJson(new { success = false, message = $"Failed to hide sheet '{settings.SheetName}': {ex.Message}" });
            return 1;
        }
    }

    private int ExecuteVeryHide(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("--sheet is required for very-hide.");
            return -1;
        }

        try
        {
            _sheetCommands.VeryHide(batch, settings.SheetName);
            _console.WriteJson(new { success = true, message = $"Sheet '{settings.SheetName}' very hidden successfully" });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteJson(new { success = false, message = $"Failed to very-hide sheet '{settings.SheetName}': {ex.Message}" });
            return 1;
        }
    }

    private static bool TryParseVisibility(string? value, out SheetVisibility visibility)
    {
        visibility = SheetVisibility.Visible;
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return value.Trim().ToLowerInvariant() switch
        {
            "visible" => SetVisibility(SheetVisibility.Visible, out visibility),
            "hidden" => SetVisibility(SheetVisibility.Hidden, out visibility),
            "very-hidden" => SetVisibility(SheetVisibility.VeryHidden, out visibility),
            "veryhidden" => SetVisibility(SheetVisibility.VeryHidden, out visibility),
            _ => false
        };
    }

    private static bool SetVisibility(SheetVisibility value, out SheetVisibility result)
    {
        result = value;
        return true;
    }

    private int WriteResult(ResultBase result)
    {
        _console.WriteJson(result);
        return result.Success ? 0 : -1;
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown sheet action '{action}'.");
        return -1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--sheet <NAME>")]
        public string? SheetName { get; init; }

        [CommandOption("--new-name <NAME>")]
        public string? NewSheetName { get; init; }

        [CommandOption("--source-sheet <NAME>")]
        public string? SourceSheet { get; init; }

        [CommandOption("--target-sheet <NAME>")]
        public string? TargetSheet { get; init; }

        [CommandOption("--red <VALUE>")]
        public int? Red { get; init; }

        [CommandOption("--green <VALUE>")]
        public int? Green { get; init; }

        [CommandOption("--blue <VALUE>")]
        public int? Blue { get; init; }

        [CommandOption("--visibility <visible|hidden|very-hidden>")]
        public string? Visibility { get; init; }
    }
}
