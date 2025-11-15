using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Worksheet lifecycle management commands - wraps Core with CLI formatting
/// Data operations (read, write, clear, append) moved to RangeCommands.
/// </summary>
public class SheetCommands : ISheetCommands
{
    private readonly Core.Commands.SheetCommands _coreCommands = new();

    public int List(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-list <file.xlsx> [[--session-id <id>]]");
            return 1;
        }

        var filePath = args[1];
        AnsiConsole.MarkupLine($"[bold]Worksheets in:[/] {Path.GetFileName(filePath)}\n");

        // Use CommandHelper to support both batch and non-batch mode
        try
        {
            var result = CommandHelper.WithBatchAsync(args, filePath, save: false,
                _coreCommands.ListAsync);

            if (result.Success)
            {
                if (result.Worksheets.Count > 0)
                {
                    var table = new Table();
                    table.AddColumn("[bold]Index[/]");
                    table.AddColumn("[bold]Worksheet Name[/]");

                    foreach (var sheet in result.Worksheets)
                    {
                        table.AddRow(sheet.Index.ToString(System.Globalization.CultureInfo.InvariantCulture), sheet.Name.EscapeMarkup());
                    }

                    AnsiConsole.Write(table);
                    AnsiConsole.MarkupLine($"\n[dim]Found {result.Worksheets.Count} worksheet(s)[/]");
                }
                else
                {
                    AnsiConsole.MarkupLine("[yellow]No worksheets found[/]");
                }
                return 0;
            }
            else
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
                return 1;
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }

    public int Create(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-create <file.xlsx> <sheet-name> [[--session-id <id>]]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        try
        {
            var result = CommandHelper.WithBatchAsync(args, filePath, save: true,
                async (batch) => await _coreCommands.CreateAsync(batch, sheetName));

            if (result.Success)
            {
                AnsiConsole.MarkupLine($"[green]✓[/] Created worksheet '{sheetName.EscapeMarkup()}'");
                return 0;
            }
            else
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
                return 1;
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }

    public int Rename(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-rename <file.xlsx> <old-name> <new-name>");
            return 1;
        }

        var filePath = args[1];
        var oldName = args[2];
        var newName = args[3];

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var renameResult = await _coreCommands.RenameAsync(batch, oldName, newName);
                await batch.SaveAsync();
                return renameResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Renamed '{oldName.EscapeMarkup()}' to '{newName.EscapeMarkup()}'");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Copy(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-copy <file.xlsx> <source-name> <target-name>");
            return 1;
        }

        var filePath = args[1];
        var sourceName = args[2];
        var targetName = args[3];

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var copyResult = await _coreCommands.CopyAsync(batch, sourceName, targetName);
                await batch.SaveAsync();
                return copyResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Copied '{sourceName.EscapeMarkup()}' to '{targetName.EscapeMarkup()}'");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Delete(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-delete <file.xlsx> <sheet-name>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var deleteResult = await _coreCommands.DeleteAsync(batch, sheetName);
                await batch.SaveAsync();
                return deleteResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Deleted worksheet '{sheetName.EscapeMarkup()}'");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === TAB COLOR COMMANDS ===

    public int SetTabColor(string[] args)
    {
        if (args.Length < 6)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-set-tab-color <file.xlsx> <sheet-name> <red> <green> <blue>");
            AnsiConsole.MarkupLine("[dim]RGB values: 0-255[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        if (!int.TryParse(args[3], out int red) || !int.TryParse(args[4], out int green) || !int.TryParse(args[5], out int blue))
        {
            AnsiConsole.MarkupLine("[red]Error:[/] RGB values must be integers (0-255)");
            return 1;
        }

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var setResult = await _coreCommands.SetTabColorAsync(batch, sheetName, red, green, blue);
                await batch.SaveAsync();
                return setResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            var hexColor = $"#{red:X2}{green:X2}{blue:X2}";
            AnsiConsole.MarkupLine($"[green]✓[/] Set tab color for '{sheetName.EscapeMarkup()}' to {hexColor} (RGB: {red}, {green}, {blue})");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int GetTabColor(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-get-tab-color <file.xlsx> <sheet-name>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        TabColorResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                return await _coreCommands.GetTabColorAsync(batch, sheetName);
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[bold]Sheet:[/] {sheetName.EscapeMarkup()}");

            if (result.HasColor)
            {
                AnsiConsole.MarkupLine($"[bold]Color:[/] {result.HexColor} (Red: {result.Red}, Green: {result.Green}, Blue: {result.Blue})");
            }
            else
            {
                AnsiConsole.MarkupLine("[dim]No custom tab color set (using default)[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int ClearTabColor(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-clear-tab-color <file.xlsx> <sheet-name>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var clearResult = await _coreCommands.ClearTabColorAsync(batch, sheetName);
                await batch.SaveAsync();
                return clearResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Cleared tab color for '{sheetName.EscapeMarkup()}'");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === VISIBILITY COMMANDS ===

    public int SetVisibility(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-set-visibility <file.xlsx> <sheet-name> <visible|hidden|veryhidden>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var visibilityStr = args[3].ToLowerInvariant();

        SheetVisibility visibility = visibilityStr switch
        {
            "visible" => SheetVisibility.Visible,
            "hidden" => SheetVisibility.Hidden,
            "veryhidden" => SheetVisibility.VeryHidden,
            _ => (SheetVisibility)(-999) // Invalid value
        };

        if ((int)visibility == -999)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Visibility must be 'visible', 'hidden', or 'veryhidden'");
            return 1;
        }

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var setResult = await _coreCommands.SetVisibilityAsync(batch, sheetName, visibility);
                await batch.SaveAsync();
                return setResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Set visibility for '{sheetName.EscapeMarkup()}' to {visibilityStr}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int GetVisibility(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-get-visibility <file.xlsx> <sheet-name>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        SheetVisibilityResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                return await _coreCommands.GetVisibilityAsync(batch, sheetName);
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[bold]Sheet:[/] {sheetName.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[bold]Visibility:[/] {result.VisibilityName}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Show(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-show <file.xlsx> <sheet-name>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var showResult = await _coreCommands.ShowAsync(batch, sheetName);
                await batch.SaveAsync();
                return showResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] '{sheetName.EscapeMarkup()}' is now visible");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Hide(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-hide <file.xlsx> <sheet-name>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var hideResult = await _coreCommands.HideAsync(batch, sheetName);
                await batch.SaveAsync();
                return hideResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] '{sheetName.EscapeMarkup()}' is now hidden (user can unhide via Excel UI)");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int VeryHide(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-very-hide <file.xlsx> <sheet-name>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var veryHideResult = await _coreCommands.VeryHideAsync(batch, sheetName);
                await batch.SaveAsync();
                return veryHideResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] '{sheetName.EscapeMarkup()}' is now very hidden (requires code to unhide)");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }
}
