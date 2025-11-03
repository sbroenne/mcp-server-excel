using Sbroenne.ExcelMcp.ComInterop.Session;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Connection management commands - wraps Core with CLI formatting
/// </summary>
public class ConnectionCommands : IConnectionCommands
{
    private readonly Core.Commands.ConnectionCommands _coreCommands = new();

    public int List(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] conn-list <file.xlsx>");
            return 1;
        }

        var filePath = args[1];
        AnsiConsole.MarkupLine($"[bold]Connections in:[/] {Path.GetFileName(filePath)}\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ListAsync(batch);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            if (result.Connections.Count > 0)
            {
                var table = new Table();
                table.AddColumn("[bold]Name[/]");
                table.AddColumn("[bold]Type[/]");
                table.AddColumn("[bold]Description[/]");
                table.AddColumn("[bold]Last Refresh[/]");
                table.AddColumn("[bold]Power Query[/]");

                foreach (var conn in result.Connections.OrderBy(c => c.Name))
                {
                    string description = conn.Description?.Length > 30 ? conn.Description[..27] + "..." : conn.Description ?? "";
                    string lastRefresh = conn.LastRefresh?.ToString("yyyy-MM-dd HH:mm") ?? "-";
                    string isPQ = conn.IsPowerQuery ? "[green]✓[/]" : "";

                    table.AddRow(
                        conn.Name.EscapeMarkup(),
                        conn.Type.EscapeMarkup(),
                        description.EscapeMarkup(),
                        lastRefresh,
                        isPQ
                    );
                }

                AnsiConsole.Write(table);
                AnsiConsole.MarkupLine($"\n[dim]Found {result.Connections.Count} connection(s)[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No connections found in this workbook[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int View(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] conn-view <file.xlsx> <connection-name>");
            return 1;
        }

        var filePath = args[1];
        var connectionName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ViewAsync(batch, connectionName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[bold cyan]Connection:[/] {result.ConnectionName.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[bold]Type:[/] {result.Type.EscapeMarkup()}");

            if (result.IsPowerQuery)
            {
                AnsiConsole.MarkupLine($"[yellow]Power Query:[/] Yes - Use pq-* commands for modifications");
            }

            if (!string.IsNullOrEmpty(result.ConnectionString))
            {
                AnsiConsole.MarkupLine($"\n[bold]Connection String:[/]");
                AnsiConsole.MarkupLine($"[dim]{result.ConnectionString.EscapeMarkup()}[/]");
            }

            if (!string.IsNullOrEmpty(result.CommandText))
            {
                AnsiConsole.MarkupLine($"\n[bold]Command Text:[/]");
                AnsiConsole.MarkupLine($"[dim]{result.CommandText.EscapeMarkup()}[/]");
            }

            if (!string.IsNullOrEmpty(result.CommandType))
            {
                AnsiConsole.MarkupLine($"[bold]Command Type:[/] {result.CommandType}");
            }

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Import(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] conn-import <file.xlsx> <connection-name> <definition.json>");
            return 1;
        }

        var filePath = args[1];
        var connectionName = args[2];
        var jsonPath = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.ImportAsync(batch, connectionName, jsonPath);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Imported connection '{connectionName.EscapeMarkup()}'");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Export(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] conn-export <file.xlsx> <connection-name> <output.json>");
            return 1;
        }

        var filePath = args[1];
        var connectionName = args[2];
        var jsonPath = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ExportAsync(batch, connectionName, jsonPath);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Exported connection '{connectionName.EscapeMarkup()}' to {Path.GetFileName(jsonPath)}");
            AnsiConsole.MarkupLine($"[dim]Output: {jsonPath}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Update(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] conn-update <file.xlsx> <connection-name> <definition.json>");
            return 1;
        }

        var filePath = args[1];
        var connectionName = args[2];
        var jsonPath = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.UpdatePropertiesAsync(batch, connectionName, jsonPath);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Updated connection '{connectionName.EscapeMarkup()}'");
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
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] conn-refresh <file.xlsx> <connection-name>");
            return 1;
        }

        var filePath = args[1];
        var connectionName = args[2];

        AnsiConsole.Status()
            .Start("Refreshing connection...", ctx =>
            {
                ctx.Spinner(Spinner.Known.Dots);
                ctx.SpinnerStyle(Style.Parse("green"));

                var task = Task.Run(async () =>
                {
                    await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                    var result = await _coreCommands.RefreshAsync(batch, connectionName);
                    await batch.SaveAsync();
                    return result;
                });
                var result = task.GetAwaiter().GetResult();

                if (result.Success)
                {
                    ctx.Status("[green]✓ Refresh complete[/]");
                }
            });

        var finalTask = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.RefreshAsync(batch, connectionName);
            await batch.SaveAsync();
            return result;
        });
        var finalResult = finalTask.GetAwaiter().GetResult();

        if (finalResult.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Refreshed connection '{connectionName.EscapeMarkup()}'");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {finalResult.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Delete(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] conn-delete <file.xlsx> <connection-name>");
            return 1;
        }

        var filePath = args[1];
        var connectionName = args[2];

        // Confirm deletion
        if (!AnsiConsole.Confirm($"Delete connection '{connectionName}'?"))
        {
            AnsiConsole.MarkupLine("[yellow]Cancelled[/]");
            return 0;
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.DeleteAsync(batch, connectionName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Deleted connection '{connectionName.EscapeMarkup()}'");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int LoadTo(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] conn-loadto <file.xlsx> <connection-name> <sheet-name>");
            return 1;
        }

        var filePath = args[1];
        var connectionName = args[2];
        var sheetName = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.LoadToAsync(batch, connectionName, sheetName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Loaded connection '{connectionName.EscapeMarkup()}' to sheet '{sheetName.EscapeMarkup()}'");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int GetProperties(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] conn-properties <file.xlsx> <connection-name>");
            return 1;
        }

        var filePath = args[1];
        var connectionName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetPropertiesAsync(batch, connectionName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[bold]Properties for:[/] {result.ConnectionName.EscapeMarkup()}\n");

            var table = new Table();
            table.AddColumn("[bold]Property[/]");
            table.AddColumn("[bold]Value[/]");

            table.AddRow("Background Query", result.BackgroundQuery ? "[green]Yes[/]" : "[dim]No[/]");
            table.AddRow("Refresh on File Open", result.RefreshOnFileOpen ? "[green]Yes[/]" : "[dim]No[/]");
            table.AddRow("Save Password", result.SavePassword ? "[yellow]Yes[/]" : "[dim]No[/]");
            table.AddRow("Refresh Period (minutes)", result.RefreshPeriod.ToString());

            AnsiConsole.Write(table);
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetProperties(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] conn-set-properties <file.xlsx> <connection-name> [--bg-query true|false] [--refresh-on-open true|false] [--save-password true|false] [--refresh-period <minutes>]");
            AnsiConsole.MarkupLine("[yellow]Example:[/] conn-set-properties data.xlsx MyConnection --bg-query true --refresh-period 30");
            return 1;
        }

        var filePath = args[1];
        var connectionName = args[2];

        // Parse optional properties
        bool? backgroundQuery = null;
        bool? refreshOnFileOpen = null;
        bool? savePassword = null;
        int? refreshPeriod = null;

        for (int i = 3; i < args.Length; i += 2)
        {
            if (i + 1 >= args.Length) break;

            string flag = args[i].ToLower();
            string value = args[i + 1];

            switch (flag)
            {
                case "--bg-query" or "--background-query":
                    if (bool.TryParse(value, out bool bq))
                        backgroundQuery = bq;
                    break;
                case "--refresh-on-open":
                    if (bool.TryParse(value, out bool ro))
                        refreshOnFileOpen = ro;
                    break;
                case "--save-password" or "--save-pwd":
                    if (bool.TryParse(value, out bool sp))
                        savePassword = sp;
                    break;
                case "--refresh-period":
                    if (int.TryParse(value, out int rp))
                        refreshPeriod = rp;
                    break;
            }
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.SetPropertiesAsync(batch, connectionName, backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Updated properties for connection '{connectionName.EscapeMarkup()}'");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Test(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] conn-test <file.xlsx> <connection-name>");
            return 1;
        }

        var filePath = args[1];
        var connectionName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.TestAsync(batch, connectionName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Connection '{connectionName.EscapeMarkup()}' is valid");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }
}
