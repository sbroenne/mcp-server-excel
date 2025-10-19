using Spectre.Console;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands implementation
/// </summary>
public class ParameterCommands : IParameterCommands
{
    public int List(string[] args)
    {
        if (!ValidateArgs(args, 2, "param-list <file.xlsx>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[bold]Named Ranges/Parameters in:[/] {Path.GetFileName(args[1])}\n");

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            var names = new List<(string Name, string RefersTo)>();

            // Get Named Ranges
            try
            {
                dynamic namesCollection = workbook.Names;
                int count = namesCollection.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic nameObj = namesCollection.Item(i);
                    string name = nameObj.Name;
                    string refersTo = nameObj.RefersTo ?? "";
                    names.Add((name, refersTo.Length > 80 ? refersTo[..77] + "..." : refersTo));
                }
            }
            catch { }

            // Display named ranges
            if (names.Count > 0)
            {
                var table = new Table();
                table.AddColumn("[bold]Parameter Name[/]");
                table.AddColumn("[bold]Value/Formula[/]");

                foreach (var (name, refersTo) in names.OrderBy(n => n.Name))
                {
                    table.AddRow(
                        $"[yellow]{name.EscapeMarkup()}[/]",
                        $"[dim]{refersTo.EscapeMarkup()}[/]"
                    );
                }

                AnsiConsole.Write(table);
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine($"[bold]Total:[/] {names.Count} named ranges");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No named ranges found[/]");
            }

            return 0;
        });
    }

    public int Set(string[] args)
    {
        if (!ValidateArgs(args, 4, "param-set <file.xlsx> <param-name> <value>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var paramName = args[2];
        var value = args[3];

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            dynamic? nameObj = FindName(workbook, paramName);
            if (nameObj == null)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Parameter '{paramName}' not found");
                return 1;
            }

            nameObj.RefersTo = value;
            workbook.Save();
            AnsiConsole.MarkupLine($"[green]✓[/] Set parameter '{paramName}' = '{value}'");
            return 0;
        });
    }

    public int Get(string[] args)
    {
        if (!ValidateArgs(args, 3, "param-get <file.xlsx> <param-name>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var paramName = args[2];

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            try
            {
                dynamic? nameObj = FindName(workbook, paramName);
                if (nameObj == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Parameter '{paramName}' not found");
                    return 1;
                }

                string refersTo = nameObj.RefersTo ?? "";
                
                // Try to get the actual value if it's a cell reference
                try
                {
                    dynamic refersToRange = nameObj.RefersToRange;
                    if (refersToRange != null)
                    {
                        object cellValue = refersToRange.Value2;
                        AnsiConsole.MarkupLine($"[cyan]{paramName}:[/] {cellValue?.ToString()?.EscapeMarkup() ?? "[null]"}");
                        AnsiConsole.MarkupLine($"[dim]Refers to: {refersTo.EscapeMarkup()}[/]");
                    }
                    else
                    {
                        AnsiConsole.MarkupLine($"[cyan]{paramName}:[/] {refersTo.EscapeMarkup()}");
                    }
                }
                catch
                {
                    // If we can't get the range value, just show the formula
                    AnsiConsole.MarkupLine($"[cyan]{paramName}:[/] {refersTo.EscapeMarkup()}");
                }

                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    public int Create(string[] args)
    {
        if (!ValidateArgs(args, 4, "param-create <file.xlsx> <param-name> <value-or-reference>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var paramName = args[2];
        var valueOrRef = args[3];

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            try
            {
                // Check if parameter already exists
                dynamic? existingName = FindName(workbook, paramName);
                if (existingName != null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Parameter '{paramName}' already exists");
                    return 1;
                }

                // Create new named range
                dynamic names = workbook.Names;
                names.Add(paramName, valueOrRef);

                workbook.Save();
                AnsiConsole.MarkupLine($"[green]✓[/] Created parameter '{paramName}' = '{valueOrRef.EscapeMarkup()}'");
                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    public int Delete(string[] args)
    {
        if (!ValidateArgs(args, 3, "param-delete <file.xlsx> <param-name>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var paramName = args[2];

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            try
            {
                dynamic? nameObj = FindName(workbook, paramName);
                if (nameObj == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Parameter '{paramName}' not found");
                    return 1;
                }

                nameObj.Delete();
                workbook.Save();
                AnsiConsole.MarkupLine($"[green]✓[/] Deleted parameter '{paramName}'");
                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }
}
