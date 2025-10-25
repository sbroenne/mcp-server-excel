using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Individual cell operation commands - wraps Core with CLI formatting
/// </summary>
public class CellCommands : ICellCommands
{
    private readonly Core.Commands.CellCommands _coreCommands = new();

    public int GetValue(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-get-value <file.xlsx> <sheet-name> <cell-address>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];

        var result = _coreCommands.GetValue(filePath, sheetName, cellAddress);

        if (result.Success)
        {
            string displayValue = result.Value?.ToString() ?? "[null]";
            AnsiConsole.MarkupLine($"[cyan]{result.CellAddress}:[/] {displayValue.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetValue(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-set-value <file.xlsx> <sheet-name> <cell-address> <value>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];
        var value = args[4];

        var result = _coreCommands.SetValue(filePath, sheetName, cellAddress, value);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Set {sheetName}!{cellAddress} = '{value.EscapeMarkup()}'");

            // Display workflow hints if available
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggested Next Actions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int GetFormula(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-get-formula <file.xlsx> <sheet-name> <cell-address>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];

        var result = _coreCommands.GetFormula(filePath, sheetName, cellAddress);

        if (result.Success)
        {
            string displayValue = result.Value?.ToString() ?? "[null]";

            if (string.IsNullOrEmpty(result.Formula))
            {
                AnsiConsole.MarkupLine($"[cyan]{result.CellAddress}:[/] [yellow](no formula)[/] Value: {displayValue.EscapeMarkup()}");
            }
            else
            {
                AnsiConsole.MarkupLine($"[cyan]{result.CellAddress}:[/] {result.Formula.EscapeMarkup()}");
                AnsiConsole.MarkupLine($"[dim]Result: {displayValue.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetFormula(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-set-formula <file.xlsx> <sheet-name> <cell-address> <formula>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];
        var formula = args[4];

        var result = _coreCommands.SetFormula(filePath, sheetName, cellAddress, formula);

        if (result.Success)
        {
            // Need to get the result value by calling GetValue
            var valueResult = _coreCommands.GetValue(filePath, sheetName, cellAddress);
            string displayResult = valueResult.Value?.ToString() ?? "[null]";

            AnsiConsole.MarkupLine($"[green]✓[/] Set {sheetName}!{cellAddress} = {formula.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[dim]Result: {displayResult.EscapeMarkup()}[/]");

            // Display workflow hints if available
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggested Next Actions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetBackgroundColor(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-set-background-color <file.xlsx> <sheet-name> <cell-address> <color>");
            AnsiConsole.MarkupLine("[dim]Color formats: #RRGGBB (hex), r,g,b (RGB), or color number[/]");
            AnsiConsole.MarkupLine("[dim]Examples: #FF0000, 255,0,0, 255[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];
        var color = args[4];

        var result = _coreCommands.SetBackgroundColor(filePath, sheetName, cellAddress, color);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Background color set for {sheetName}!{cellAddress}");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetFontColor(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-set-font-color <file.xlsx> <sheet-name> <cell-address> <color>");
            AnsiConsole.MarkupLine("[dim]Color formats: #RRGGBB (hex), r,g,b (RGB), or color number[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];
        var color = args[4];

        var result = _coreCommands.SetFontColor(filePath, sheetName, cellAddress, color);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Font color set for {sheetName}!{cellAddress}");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetFont(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-set-font <file.xlsx> <sheet-name> <cell-address> [name=Arial] [size=11] [bold=true|false] [italic=true|false] [underline=true|false]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];

        string? fontName = null;
        int? fontSize = null;
        bool? bold = null;
        bool? italic = null;
        bool? underline = null;

        // Parse optional named parameters
        for (int i = 4; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg.Contains("="))
            {
                var parts = arg.Split('=', 2);
                var key = parts[0].ToLowerInvariant();
                var value = parts[1];

                switch (key)
                {
                    case "name":
                        fontName = value;
                        break;
                    case "size":
                        if (int.TryParse(value, out int size)) fontSize = size;
                        break;
                    case "bold":
                        if (bool.TryParse(value, out bool b)) bold = b;
                        break;
                    case "italic":
                        if (bool.TryParse(value, out bool it)) italic = it;
                        break;
                    case "underline":
                        if (bool.TryParse(value, out bool u)) underline = u;
                        break;
                }
            }
        }

        var result = _coreCommands.SetFont(filePath, sheetName, cellAddress, fontName, fontSize, bold, italic, underline);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Font properties set for {sheetName}!{cellAddress}");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetBorder(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-set-border <file.xlsx> <sheet-name> <cell-address> <style> [color]");
            AnsiConsole.MarkupLine("[dim]Styles: thin, dash, dot, double, none[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];
        var borderStyle = args[4];
        var borderColor = args.Length > 5 ? args[5] : null;

        var result = _coreCommands.SetBorder(filePath, sheetName, cellAddress, borderStyle, borderColor);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Border set for {sheetName}!{cellAddress}");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetNumberFormat(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-set-number-format <file.xlsx> <sheet-name> <cell-address> <format>");
            AnsiConsole.MarkupLine("[dim]Examples: $#,##0.00 (currency), 0.00% (percentage), m/d/yyyy (date)[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];
        var format = args[4];

        var result = _coreCommands.SetNumberFormat(filePath, sheetName, cellAddress, format);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Number format set for {sheetName}!{cellAddress}");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetAlignment(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-set-alignment <file.xlsx> <sheet-name> <cell-address> [horizontal=left|center|right|justify] [vertical=top|center|bottom] [wrap=true|false]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];

        string? horizontal = null;
        string? vertical = null;
        bool? wrapText = null;

        // Parse optional named parameters
        for (int i = 4; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg.Contains("="))
            {
                var parts = arg.Split('=', 2);
                var key = parts[0].ToLowerInvariant();
                var value = parts[1];

                switch (key)
                {
                    case "horizontal":
                        horizontal = value;
                        break;
                    case "vertical":
                        vertical = value;
                        break;
                    case "wrap":
                        if (bool.TryParse(value, out bool w)) wrapText = w;
                        break;
                }
            }
        }

        var result = _coreCommands.SetAlignment(filePath, sheetName, cellAddress, horizontal, vertical, wrapText);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Alignment set for {sheetName}!{cellAddress}");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int ClearFormatting(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-clear-formatting <file.xlsx> <sheet-name> <cell-address>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];

        var result = _coreCommands.ClearFormatting(filePath, sheetName, cellAddress);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Formatting cleared for {sheetName}!{cellAddress}");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }
}
