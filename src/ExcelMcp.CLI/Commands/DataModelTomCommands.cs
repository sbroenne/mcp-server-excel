using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Data Model TOM (Tabular Object Model) commands - wraps Core with CLI formatting
/// </summary>
public class DataModelTomCommands : IDataModelTomCommands
{
    private readonly Core.Commands.DataModelTomCommands _coreCommands = new();

    public int CreateMeasure(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-create-measure <file.xlsx> <table-name> <measure-name> <dax-formula> [--desc <description>] [--format <format-string>]");
            AnsiConsole.MarkupLine("\n[bold]Examples:[/]");
            AnsiConsole.MarkupLine("  dm-create-measure Sales.xlsx Sales \"Total Sales\" \"SUM(Sales[Amount])\"");
            AnsiConsole.MarkupLine("  dm-create-measure Sales.xlsx Sales \"Avg Price\" \"AVERAGE(Sales[Price])\" --format \"#,##0.00\"");
            return 1;
        }

        var filePath = args[1];
        var tableName = args[2];
        var measureName = args[3];
        var daxFormula = args[4];

        // Parse optional parameters
        string? description = null;
        string? formatString = null;

        for (int i = 5; i < args.Length; i++)
        {
            if (args[i] == "--desc" && i + 1 < args.Length)
            {
                description = args[i + 1];
                i++;
            }
            else if (args[i] == "--format" && i + 1 < args.Length)
            {
                formatString = args[i + 1];
                i++;
            }
        }

        AnsiConsole.Status()
            .Start($"Creating measure [bold]{measureName.EscapeMarkup()}[/] in table [bold]{tableName.EscapeMarkup()}[/]...", ctx =>
            {
                ctx.Spinner(Spinner.Known.Dots);
                ctx.SpinnerStyle(Style.Parse("green"));
                System.Threading.Thread.Sleep(100); // Brief pause for visual feedback
            });

        var result = _coreCommands.CreateMeasure(
            filePath,
            tableName,
            measureName,
            daxFormula,
            description,
            formatString
        );

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Measure [bold]{measureName.EscapeMarkup()}[/] created successfully");

            var panel = new Panel($"[bold]Table:[/] {tableName.EscapeMarkup()}\n" +
                                 $"[bold]Measure:[/] {measureName.EscapeMarkup()}\n" +
                                 $"[bold]Formula:[/] {daxFormula.EscapeMarkup()}")
            {
                Header = new PanelHeader("Measure Details"),
                Border = BoxBorder.Rounded
            };
            AnsiConsole.Write(panel);

            // Display suggested next actions
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
            AnsiConsole.MarkupLine($"[red]✗ Error:[/] {result.ErrorMessage?.EscapeMarkup()}");

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggestions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 1;
        }
    }

    public int UpdateMeasure(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-update-measure <file.xlsx> <measure-name> [--formula <dax-formula>] [--desc <description>] [--format <format-string>]");
            AnsiConsole.MarkupLine("\n[bold]Examples:[/]");
            AnsiConsole.MarkupLine("  dm-update-measure Sales.xlsx \"Total Sales\" --formula \"SUM(Sales[Amount]) * 1.1\"");
            AnsiConsole.MarkupLine("  dm-update-measure Sales.xlsx \"Avg Price\" --format \"#,##0.00\" --desc \"Updated description\"");
            return 1;
        }

        var filePath = args[1];
        var measureName = args[2];

        // Parse optional parameters
        string? daxFormula = null;
        string? description = null;
        string? formatString = null;

        for (int i = 3; i < args.Length; i++)
        {
            if (args[i] == "--formula" && i + 1 < args.Length)
            {
                daxFormula = args[i + 1];
                i++;
            }
            else if (args[i] == "--desc" && i + 1 < args.Length)
            {
                description = args[i + 1];
                i++;
            }
            else if (args[i] == "--format" && i + 1 < args.Length)
            {
                formatString = args[i + 1];
                i++;
            }
        }

        AnsiConsole.Status()
            .Start($"Updating measure [bold]{measureName.EscapeMarkup()}[/]...", ctx =>
            {
                ctx.Spinner(Spinner.Known.Dots);
                ctx.SpinnerStyle(Style.Parse("green"));
                System.Threading.Thread.Sleep(100);
            });

        var result = _coreCommands.UpdateMeasure(
            filePath,
            measureName,
            daxFormula,
            description,
            formatString
        );

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Measure [bold]{measureName.EscapeMarkup()}[/] updated successfully");

            // Display suggested next actions
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
            AnsiConsole.MarkupLine($"[red]✗ Error:[/] {result.ErrorMessage?.EscapeMarkup()}");

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggestions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 1;
        }
    }

    public int CreateRelationship(string[] args)
    {
        if (args.Length < 6)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-create-relationship <file.xlsx> <from-table> <from-column> <to-table> <to-column> [--inactive] [--bidirectional]");
            AnsiConsole.MarkupLine("\n[bold]Examples:[/]");
            AnsiConsole.MarkupLine("  dm-create-relationship Sales.xlsx Sales CustomerID Customers CustomerID");
            AnsiConsole.MarkupLine("  dm-create-relationship Sales.xlsx Sales ProductID Products ProductID --bidirectional");
            return 1;
        }

        var filePath = args[1];
        var fromTable = args[2];
        var fromColumn = args[3];
        var toTable = args[4];
        var toColumn = args[5];

        // Parse optional flags
        bool isActive = true;
        string crossFilterDirection = "Single";

        for (int i = 6; i < args.Length; i++)
        {
            if (args[i] == "--inactive")
            {
                isActive = false;
            }
            else if (args[i] == "--bidirectional" || args[i] == "--both")
            {
                crossFilterDirection = "Both";
            }
        }

        AnsiConsole.Status()
            .Start($"Creating relationship from [bold]{fromTable}.{fromColumn}[/] to [bold]{toTable}.{toColumn}[/]...", ctx =>
            {
                ctx.Spinner(Spinner.Known.Dots);
                ctx.SpinnerStyle(Style.Parse("green"));
                System.Threading.Thread.Sleep(100);
            });

        var result = _coreCommands.CreateRelationship(
            filePath,
            fromTable,
            fromColumn,
            toTable,
            toColumn,
            isActive,
            crossFilterDirection
        );

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Relationship created successfully");

            var panel = new Panel($"[bold]From:[/] {fromTable.EscapeMarkup()}.{fromColumn.EscapeMarkup()}\n" +
                                 $"[bold]To:[/] {toTable.EscapeMarkup()}.{toColumn.EscapeMarkup()}\n" +
                                 $"[bold]Active:[/] {(isActive ? "Yes" : "No")}\n" +
                                 $"[bold]Cross-Filter:[/] {crossFilterDirection}")
            {
                Header = new PanelHeader("Relationship Details"),
                Border = BoxBorder.Rounded
            };
            AnsiConsole.Write(panel);

            // Display suggested next actions
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
            AnsiConsole.MarkupLine($"[red]✗ Error:[/] {result.ErrorMessage?.EscapeMarkup()}");

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggestions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 1;
        }
    }

    public int UpdateRelationship(string[] args)
    {
        if (args.Length < 6)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-update-relationship <file.xlsx> <from-table> <from-column> <to-table> <to-column> [--active|--inactive] [--single|--bidirectional]");
            AnsiConsole.MarkupLine("\n[bold]Examples:[/]");
            AnsiConsole.MarkupLine("  dm-update-relationship Sales.xlsx Sales CustomerID Customers CustomerID --inactive");
            AnsiConsole.MarkupLine("  dm-update-relationship Sales.xlsx Sales ProductID Products ProductID --bidirectional");
            return 1;
        }

        var filePath = args[1];
        var fromTable = args[2];
        var fromColumn = args[3];
        var toTable = args[4];
        var toColumn = args[5];

        // Parse optional flags
        bool? isActive = null;
        string? crossFilterDirection = null;

        for (int i = 6; i < args.Length; i++)
        {
            if (args[i] == "--active")
            {
                isActive = true;
            }
            else if (args[i] == "--inactive")
            {
                isActive = false;
            }
            else if (args[i] == "--single")
            {
                crossFilterDirection = "Single";
            }
            else if (args[i] == "--bidirectional" || args[i] == "--both")
            {
                crossFilterDirection = "Both";
            }
        }

        AnsiConsole.Status()
            .Start($"Updating relationship...", ctx =>
            {
                ctx.Spinner(Spinner.Known.Dots);
                ctx.SpinnerStyle(Style.Parse("green"));
                System.Threading.Thread.Sleep(100);
            });

        var result = _coreCommands.UpdateRelationship(
            filePath,
            fromTable,
            fromColumn,
            toTable,
            toColumn,
            isActive,
            crossFilterDirection
        );

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Relationship updated successfully");

            // Display suggested next actions
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
            AnsiConsole.MarkupLine($"[red]✗ Error:[/] {result.ErrorMessage?.EscapeMarkup()}");

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggestions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 1;
        }
    }

    public int CreateCalculatedColumn(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-create-column <file.xlsx> <table-name> <column-name> <dax-formula> [--type <data-type>] [--desc <description>]");
            AnsiConsole.MarkupLine("\n[bold]Data Types:[/] String, Integer, Double, Boolean, DateTime");
            AnsiConsole.MarkupLine("\n[bold]Examples:[/]");
            AnsiConsole.MarkupLine("  dm-create-column Sales.xlsx Sales TotalCost \"[Price] * [Quantity]\" --type Double");
            AnsiConsole.MarkupLine("  dm-create-column Sales.xlsx Sales IsHighValue \"[Amount] > 1000\" --type Boolean");
            return 1;
        }

        var filePath = args[1];
        var tableName = args[2];
        var columnName = args[3];
        var daxFormula = args[4];

        // Parse optional parameters
        string dataType = "String";
        string? description = null;

        for (int i = 5; i < args.Length; i++)
        {
            if (args[i] == "--type" && i + 1 < args.Length)
            {
                dataType = args[i + 1];
                i++;
            }
            else if (args[i] == "--desc" && i + 1 < args.Length)
            {
                description = args[i + 1];
                i++;
            }
        }

        AnsiConsole.Status()
            .Start($"Creating calculated column [bold]{columnName.EscapeMarkup()}[/]...", ctx =>
            {
                ctx.Spinner(Spinner.Known.Dots);
                ctx.SpinnerStyle(Style.Parse("green"));
                System.Threading.Thread.Sleep(100);
            });

        var result = _coreCommands.CreateCalculatedColumn(
            filePath,
            tableName,
            columnName,
            daxFormula,
            description,
            dataType
        );

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Calculated column [bold]{columnName.EscapeMarkup()}[/] created successfully");

            var panel = new Panel($"[bold]Table:[/] {tableName.EscapeMarkup()}\n" +
                                 $"[bold]Column:[/] {columnName.EscapeMarkup()}\n" +
                                 $"[bold]Formula:[/] {daxFormula.EscapeMarkup()}\n" +
                                 $"[bold]Data Type:[/] {dataType}")
            {
                Header = new PanelHeader("Column Details"),
                Border = BoxBorder.Rounded
            };
            AnsiConsole.Write(panel);

            // Display suggested next actions
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
            AnsiConsole.MarkupLine($"[red]✗ Error:[/] {result.ErrorMessage?.EscapeMarkup()}");

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggestions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 1;
        }
    }

    public int ValidateDax(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-validate-dax <file.xlsx> <dax-formula>");
            AnsiConsole.MarkupLine("\n[bold]Examples:[/]");
            AnsiConsole.MarkupLine("  dm-validate-dax Sales.xlsx \"SUM(Sales[Amount])\"");
            AnsiConsole.MarkupLine("  dm-validate-dax Sales.xlsx \"CALCULATE(SUM(Sales[Amount]), Sales[Region]=\\\"North\\\")\"");
            return 1;
        }

        var filePath = args[1];
        var daxFormula = args[2];

        AnsiConsole.Status()
            .Start($"Validating DAX formula...", ctx =>
            {
                ctx.Spinner(Spinner.Known.Dots);
                ctx.SpinnerStyle(Style.Parse("green"));
                System.Threading.Thread.Sleep(100);
            });

        var result = _coreCommands.ValidateDax(filePath, daxFormula);

        if (result.Success)
        {
            if (result.IsValid)
            {
                AnsiConsole.MarkupLine($"[green]✓[/] DAX formula appears valid");

                var panel = new Panel(daxFormula.EscapeMarkup())
                {
                    Header = new PanelHeader("Validated Formula", Justify.Left),
                    Border = BoxBorder.Rounded,
                    BorderStyle = new Style(Color.Green)
                };
                AnsiConsole.Write(panel);
            }
            else
            {
                AnsiConsole.MarkupLine($"[yellow]⚠[/] DAX formula validation issues detected");
                AnsiConsole.MarkupLine($"[red]Error:[/] {result.ValidationError?.EscapeMarkup()}");

                var panel = new Panel(daxFormula.EscapeMarkup())
                {
                    Header = new PanelHeader("Invalid Formula", Justify.Left),
                    Border = BoxBorder.Rounded,
                    BorderStyle = new Style(Color.Red)
                };
                AnsiConsole.Write(panel);
            }

            // Display suggested next actions
            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggestions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return result.IsValid ? 0 : 1;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]✗ Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }
}
