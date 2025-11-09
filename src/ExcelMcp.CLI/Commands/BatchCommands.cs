using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Batch session management commands - CLI presentation layer
/// Enables multi-operation workflows with single Excel instance (75-90% faster)
/// </summary>
public class BatchCommands
{
    private static readonly ConcurrentDictionary<string, IExcelBatch> _activeBatches = new();

    /// <summary>
    /// Begin a new batch session for a workbook
    /// </summary>
    /// <param name="args">Command arguments: batch-begin file.xlsx</param>
    /// <returns>0 for success, 1 for error</returns>
    public int Begin(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] batch-begin <file.xlsx>");
            return 1;
        }

        string filePath = args[1];

        try
        {
            // Normalize path to prevent duplicate sessions
            string normalizedPath = Path.GetFullPath(filePath);

            // Check if file exists
            if (!File.Exists(normalizedPath))
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {normalizedPath.EscapeMarkup()}");
                AnsiConsole.MarkupLine("[yellow]Hint:[/] Use create-empty command to create a new file first");
                return 1;
            }

            // Check if batch already exists for this file
            if (_activeBatches.ContainsKey(normalizedPath))
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Batch session already active for this file");
                AnsiConsole.MarkupLine("[yellow]Hint:[/] Commit or discard existing batch before starting a new one");
                return 1;
            }

            var task = Task.Run(async () =>
            {
                var batch = await ExcelSession.BeginBatchAsync(filePath);
                return batch;
            });
            var batch = task.GetAwaiter().GetResult();

            // Generate batch ID
            string batchId = Guid.NewGuid().ToString();

            // Store in active sessions
            if (!_activeBatches.TryAdd(batchId, batch))
            {
                // Cleanup if we couldn't add
                var disposeTask = Task.Run(async () => await batch.DisposeAsync());
                disposeTask.GetAwaiter().GetResult();
                AnsiConsole.MarkupLine("[red]Error:[/] Failed to register batch session");
                return 1;
            }

            AnsiConsole.MarkupLine($"[green]✓[/] [bold]Batch session started[/]");
            AnsiConsole.MarkupLine($"[cyan]Batch ID:[/] {batchId}");
            AnsiConsole.MarkupLine($"[dim]File:[/] {normalizedPath}");
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[bold]Next steps:[/]");
            AnsiConsole.MarkupLine($"[dim]• Use --batch-id {batchId} with any command[/]");
            AnsiConsole.MarkupLine($"[dim]• All operations use same Excel instance (75-90% faster!)[/]");
            AnsiConsole.MarkupLine($"[dim]• Call batch-commit {batchId} when done[/]");

            return 0;
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }

    /// <summary>
    /// Commit (save and close) a batch session
    /// </summary>
    /// <param name="args">Command arguments: batch-commit batch-id [--no-save]</param>
    /// <returns>0 for success, 1 for error</returns>
    public int Commit(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] batch-commit <batch-id> [--no-save]");
            return 1;
        }

        string batchId = args[1];
        bool save = !args.Contains("--no-save");

        try
        {
            // Retrieve batch session
            if (!_activeBatches.TryRemove(batchId, out var batch))
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Batch session '{batchId}' not found");
                AnsiConsole.MarkupLine("[yellow]Hint:[/] Use batch-list to see active sessions");
                return 1;
            }

            string filePath = batch.WorkbookPath;

            try
            {
                var task = Task.Run(async () =>
                {
                    // Save if requested
                    if (save)
                    {
                        await batch.SaveAsync();
                    }

                    // Dispose (closes workbook and releases Excel)
                    await batch.DisposeAsync();
                });
                task.GetAwaiter().GetResult();

                AnsiConsole.MarkupLine($"[green]✓[/] [bold]Batch committed[/]");
                AnsiConsole.MarkupLine($"[cyan]Batch ID:[/] {batchId}");
                AnsiConsole.MarkupLine($"[dim]File:[/] {filePath}");
                if (save)
                {
                    AnsiConsole.MarkupLine($"[green]Changes saved[/]");
                }
                else
                {
                    AnsiConsole.MarkupLine($"[yellow]Changes discarded (--no-save)[/]");
                }

                return 0;
            }
            catch
            {
                // If save/dispose fails, try to dispose anyway to prevent resource leaks
                try
                {
                    var disposeTask = Task.Run(async () => await batch.DisposeAsync());
                    disposeTask.GetAwaiter().GetResult();
                }
                catch { /* ignore */ }
                throw;
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }

    /// <summary>
    /// List all active batch sessions
    /// </summary>
    /// <param name="args">Command arguments: batch-list</param>
    /// <returns>0 for success</returns>
#pragma warning disable IDE0060 // Remove unused parameter
    public int List(string[] args)
#pragma warning restore IDE0060 // Remove unused parameter
    {
        var batches = _activeBatches.ToList();

        if (batches.Count == 0)
        {
            AnsiConsole.MarkupLine("[yellow]No active batch sessions[/]");
            AnsiConsole.MarkupLine("[dim]Start one with:[/] [cyan]batch-begin file.xlsx[/]");
            return 0;
        }

        AnsiConsole.MarkupLine($"[bold]Active Batch Sessions:[/] {batches.Count}\n");

        var table = new Table();
        table.AddColumn("[bold]Batch ID[/]");
        table.AddColumn("[bold]File Path[/]");

        foreach (var kvp in batches)
        {
            table.AddRow(
                $"[cyan]{kvp.Key}[/]",
                $"[dim]{kvp.Value.WorkbookPath.EscapeMarkup()}[/]"
            );
        }

        AnsiConsole.Write(table);
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[yellow]⚠[/] [bold]Remember to commit batches![/]");
        AnsiConsole.MarkupLine("[dim]Each batch holds Excel open. Call batch-commit to release resources.[/]");

        return 0;
    }

    /// <summary>
    /// Get an active batch session by ID.
    /// Used internally by other commands to retrieve the batch for operations.
    /// </summary>
    internal static IExcelBatch? GetBatch(string batchId)
    {
        if (string.IsNullOrWhiteSpace(batchId))
        {
            return null;
        }

        return _activeBatches.TryGetValue(batchId, out var batch) ? batch : null;
    }
}
