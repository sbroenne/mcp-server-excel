using Sbroenne.ExcelMcp.ComInterop.Session;
using Spectre.Console;

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

    // Phase 2: Discovery operations

    public int ListColumns(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-list-columns <file.xlsx> <table-name>");
            return 1;
        }

        var filePath = args[1];
        var tableName = args[2];
        AnsiConsole.MarkupLine($"[bold]Columns in table '{tableName}':[/] {Path.GetFileName(filePath)}\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ListTableColumnsAsync(batch, tableName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            if (result.Columns != null && result.Columns.Count > 0)
            {
                var table = new Table();
                table.AddColumn("[bold]Column Name[/]");
                table.AddColumn("[bold]Data Type[/]");
                table.AddColumn("[bold]Calculated[/]");

                foreach (var column in result.Columns.OrderBy(c => c.Name))
                {
                    table.AddRow(
                        column.Name.EscapeMarkup(),
                        column.DataType.EscapeMarkup(),
                        column.IsCalculated ? "[green]Yes[/]" : "[dim]No[/]"
                    );
                }

                AnsiConsole.Write(table);
                AnsiConsole.MarkupLine($"\n[dim]Found {result.Columns.Count} column(s) in '{tableName}'[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No columns found[/]");
            }

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

    public int ViewTable(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-view-table <file.xlsx> <table-name>");
            return 1;
        }

        var filePath = args[1];
        var tableName = args[2];
        AnsiConsole.MarkupLine($"[bold]Table Details:[/] {tableName}\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ViewTableAsync(batch, tableName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            var infoTable = new Table();
            infoTable.AddColumn("[bold]Property[/]");
            infoTable.AddColumn("[bold]Value[/]");

            infoTable.AddRow("Table Name", result.TableName.EscapeMarkup());
            infoTable.AddRow("Source", result.SourceName?.EscapeMarkup() ?? "[dim]N/A[/]");
            infoTable.AddRow("Record Count", result.RecordCount.ToString());
            infoTable.AddRow("Column Count", result.Columns?.Count.ToString() ?? "0");
            infoTable.AddRow("Measure Count", result.MeasureCount.ToString());

            AnsiConsole.Write(infoTable);

            if (result.Columns != null && result.Columns.Count > 0)
            {
                AnsiConsole.MarkupLine("\n[bold]Columns:[/]");
                var columnTable = new Table();
                columnTable.AddColumn("[bold]Name[/]");
                columnTable.AddColumn("[bold]Type[/]");
                columnTable.AddColumn("[bold]Calculated[/]");

                foreach (var column in result.Columns.OrderBy(c => c.Name))
                {
                    columnTable.AddRow(
                        column.Name.EscapeMarkup(),
                        column.DataType.EscapeMarkup(),
                        column.IsCalculated ? "[green]Yes[/]" : "[dim]No[/]"
                    );
                }

                AnsiConsole.Write(columnTable);
            }

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

    public int GetModelInfo(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-get-model-info <file.xlsx>");
            return 1;
        }

        var filePath = args[1];
        AnsiConsole.MarkupLine($"[bold]Data Model Overview:[/] {Path.GetFileName(filePath)}\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetModelInfoAsync(batch);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            var table = new Table();
            table.AddColumn("[bold]Statistic[/]");
            table.AddColumn("[bold]Count[/]", column => column.RightAligned());

            table.AddRow("Tables", result.TableCount.ToString());
            table.AddRow("Measures", result.MeasureCount.ToString());
            table.AddRow("Relationships", result.RelationshipCount.ToString());
            table.AddRow("Total Rows", result.TotalRows.ToString("N0"));

            AnsiConsole.Write(table);

            if (result.TableNames != null && result.TableNames.Count > 0)
            {
                AnsiConsole.MarkupLine($"\n[bold]Tables:[/] {string.Join(", ", result.TableNames.Select(t => t.EscapeMarkup()))}");
            }

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

    // Phase 2: CREATE/UPDATE operations

    public int CreateMeasure(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-create-measure <file.xlsx> <table-name> <measure-name> <dax-formula> [format-type] [description]");
            AnsiConsole.MarkupLine("[dim]Format types: Currency, Decimal, Percentage, General[/]");
            return 1;
        }

        var filePath = args[1];
        var tableName = args[2];
        var measureName = args[3];
        var daxFormula = args[4];
        var formatType = args.Length > 5 ? args[5] : null;
        var description = args.Length > 6 ? args[6] : null;

        AnsiConsole.MarkupLine($"[bold]Creating measure:[/] {measureName} in table {tableName}");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.CreateMeasureAsync(batch, tableName, measureName, daxFormula, formatType, description);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Measure '{measureName}' created successfully");

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

    public int UpdateMeasure(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-update-measure <file.xlsx> <measure-name> [dax-formula] [format-type] [description]");
            AnsiConsole.MarkupLine("[dim]At least one optional parameter must be provided[/]");
            AnsiConsole.MarkupLine("[dim]Format types: Currency, Decimal, Percentage, General[/]");
            return 1;
        }

        var filePath = args[1];
        var measureName = args[2];
        var daxFormula = args.Length > 3 ? args[3] : null;
        var formatType = args.Length > 4 ? args[4] : null;
        var description = args.Length > 5 ? args[5] : null;

        AnsiConsole.MarkupLine($"[bold]Updating measure:[/] {measureName}");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.UpdateMeasureAsync(batch, measureName, daxFormula, formatType, description);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Measure '{measureName}' updated successfully");

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

    public int CreateRelationship(string[] args)
    {
        if (args.Length < 6)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-create-relationship <file.xlsx> <from-table> <from-column> <to-table> <to-column> [active:true|false]");
            return 1;
        }

        var filePath = args[1];
        var fromTable = args[2];
        var fromColumn = args[3];
        var toTable = args[4];
        var toColumn = args[5];
        var active = args.Length <= 6 || bool.Parse(args[6]); // Default to true if not specified

        AnsiConsole.MarkupLine($"[bold]Creating relationship:[/] {fromTable}.{fromColumn} → {toTable}.{toColumn}");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.CreateRelationshipAsync(batch, fromTable, fromColumn, toTable, toColumn, active);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Relationship created successfully");

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

    public int UpdateRelationship(string[] args)
    {
        if (args.Length < 7)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] dm-update-relationship <file.xlsx> <from-table> <from-column> <to-table> <to-column> <active:true|false>");
            return 1;
        }

        var filePath = args[1];
        var fromTable = args[2];
        var fromColumn = args[3];
        var toTable = args[4];
        var toColumn = args[5];
        var active = bool.Parse(args[6]);

        AnsiConsole.MarkupLine($"[bold]Updating relationship:[/] {fromTable}.{fromColumn} → {toTable}.{toColumn}");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.UpdateRelationshipAsync(batch, fromTable, fromColumn, toTable, toColumn, active);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Relationship updated successfully");

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
