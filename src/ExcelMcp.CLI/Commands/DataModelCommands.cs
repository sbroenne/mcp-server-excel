using Spectre.Console;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Data Model commands - wraps Core with CLI formatting
/// </summary>
public class DataModelCommands : IDataModelCommands
{
    private readonly Core.Commands.DataModelCommands _coreCommands = new();

    public int ListTables(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-list-tables <file.xlsx>");
            return 1;
        }

        var filePath = args[1];
        AnsiConsole.MarkupLine($"[bold]Data Model Tables in:[/] {Path.GetFileName(filePath)}\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ListTablesAsync(batch);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            if (result.Tables != null && result.Tables.Count > 0)
            {
                var table = new Table();
                table.AddColumn("[bold]Table Name[/]");
                table.AddColumn("[bold]Records[/]", column => column.RightAligned());
                table.AddColumn("[bold]Source[/]");

                foreach (var dmTable in result.Tables.OrderBy(t => t.Name))
                {
                    table.AddRow(
                        dmTable.Name.EscapeMarkup(),
                        dmTable.RecordCount.ToString(),
                        dmTable.SourceName?.EscapeMarkup() ?? "[dim]N/A[/]"
                    );
                }

                AnsiConsole.Write(table);
                AnsiConsole.MarkupLine($"\n[dim]Found {result.Tables.Count} table(s) in Data Model[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No tables found in Data Model[/]");
            }

            // Display workflow hints if available
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"\n[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
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

    public int ListMeasures(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-list-measures <file.xlsx>");
            return 1;
        }

        var filePath = args[1];
        AnsiConsole.MarkupLine($"[bold]DAX Measures in:[/] {Path.GetFileName(filePath)}\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ListMeasuresAsync(batch);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            if (result.Measures != null && result.Measures.Count > 0)
            {
                var table = new Table();
                table.AddColumn("[bold]Measure Name[/]");
                table.AddColumn("[bold]Table[/]");
                table.AddColumn("[bold]Formula Preview[/]");

                foreach (var measure in result.Measures.OrderBy(m => m.Table).ThenBy(m => m.Name))
                {
                    string formulaPreview = measure.FormulaPreview?.Length > 60
                        ? measure.FormulaPreview[..57] + "..."
                        : measure.FormulaPreview ?? "[dim]N/A[/]";

                    table.AddRow(
                        measure.Name.EscapeMarkup(),
                        measure.Table.EscapeMarkup(),
                        formulaPreview.EscapeMarkup()
                    );
                }

                AnsiConsole.Write(table);
                AnsiConsole.MarkupLine($"\n[dim]Found {result.Measures.Count} measure(s) in Data Model[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No measures found in Data Model[/]");
            }

            // Display workflow hints if available
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"\n[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
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

    public int ViewMeasure(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-view-measure <file.xlsx> <measure-name>");
            return 1;
        }

        var filePath = args[1];
        var measureName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ViewMeasureAsync(batch, measureName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[bold]Measure:[/] {measureName.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[bold]Table:[/] {result.TableName?.EscapeMarkup() ?? "[dim]N/A[/]"}\n");

            var panel = new Panel(result.DaxFormula?.EscapeMarkup() ?? "[dim]No formula[/]")
                .Header("[cyan]DAX Formula[/]")
                .Border(BoxBorder.Rounded);

            AnsiConsole.Write(panel);

            // Display workflow hints if available
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"\n[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
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

    public int ExportMeasure(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-export-measure <file.xlsx> <measure-name> <output.dax>");
            return 1;
        }

        var filePath = args[1];
        var measureName = args[2];
        var outputPath = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ExportMeasureAsync(batch, measureName, outputPath);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Exported measure '{measureName.EscapeMarkup()}' to {Path.GetFileName(outputPath).EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[dim]Full path: {Path.GetFullPath(outputPath).EscapeMarkup()}[/]");

            // Display workflow hints if available
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"\n[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
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

    public int ListRelationships(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-list-relationships <file.xlsx>");
            return 1;
        }

        var filePath = args[1];
        AnsiConsole.MarkupLine($"[bold]Data Model Relationships in:[/] {Path.GetFileName(filePath)}\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ListRelationshipsAsync(batch);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            if (result.Relationships != null && result.Relationships.Count > 0)
            {
                var table = new Table();
                table.AddColumn("[bold]From Table[/]");
                table.AddColumn("[bold]Column[/]");
                table.AddColumn("[bold]→[/]", column => column.Centered());
                table.AddColumn("[bold]To Table[/]");
                table.AddColumn("[bold]Column[/]");
                table.AddColumn("[bold]Active[/]", column => column.Centered());

                foreach (var rel in result.Relationships.OrderBy(r => r.FromTable).ThenBy(r => r.ToTable))
                {
                    string active = rel.IsActive ? "[green]✓[/]" : "[dim]○[/]";

                    table.AddRow(
                        rel.FromTable.EscapeMarkup(),
                        rel.FromColumn.EscapeMarkup(),
                        "→",
                        rel.ToTable.EscapeMarkup(),
                        rel.ToColumn.EscapeMarkup(),
                        active
                    );
                }

                AnsiConsole.Write(table);
                AnsiConsole.MarkupLine($"\n[dim]Found {result.Relationships.Count} relationship(s) in Data Model[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No relationships found in Data Model[/]");
            }

            // Display workflow hints if available
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"\n[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
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

    public int Refresh(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-refresh <file.xlsx>");
            return 1;
        }

        var filePath = args[1];

        AnsiConsole.Status()
            .Start($"Refreshing Data Model in {Path.GetFileName(filePath)}...", ctx =>
            {
                ctx.Spinner(Spinner.Known.Dots);
                ctx.SpinnerStyle(Style.Parse("cyan"));

                // Small delay to show spinner
                System.Threading.Thread.Sleep(100);
            });

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.RefreshAsync(batch);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Data Model refreshed successfully");

            // Display workflow hints if available
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"\n[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
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

    public int DeleteMeasure(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-delete-measure <file.xlsx> <measure-name>");
            return 1;
        }

        var filePath = args[1];
        var measureName = args[2];

        AnsiConsole.MarkupLine($"[bold]Deleting measure:[/] {measureName.EscapeMarkup()} from {Path.GetFileName(filePath)}");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.DeleteMeasureAsync(batch, measureName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Measure '{measureName.EscapeMarkup()}' deleted successfully");

            // Display workflow hints if available
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"\n[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
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

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[yellow]Suggestions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 1;
        }
    }

    public int DeleteRelationship(string[] args)
    {
        if (args.Length < 6)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-delete-relationship <file.xlsx> <from-table> <from-column> <to-table> <to-column>");
            return 1;
        }

        var filePath = args[1];
        var fromTable = args[2];
        var fromColumn = args[3];
        var toTable = args[4];
        var toColumn = args[5];

        AnsiConsole.MarkupLine($"[bold]Deleting relationship:[/] {fromTable.EscapeMarkup()}.{fromColumn.EscapeMarkup()} → {toTable.EscapeMarkup()}.{toColumn.EscapeMarkup()}");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.DeleteRelationshipAsync(batch, fromTable, fromColumn, toTable, toColumn);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Relationship deleted successfully");

            // Display workflow hints if available
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"\n[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
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

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[yellow]Suggestions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 1;
        }
    }
}
