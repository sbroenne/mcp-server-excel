using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Session management commands - CLI presentation layer
/// Enables multi-operation workflows with single Excel instance for fast batch processing
/// Aligned with MCP Server session pattern: open → operate → save → close
/// </summary>
public class BatchCommands
{
    private static readonly ConcurrentDictionary<string, IExcelBatch> _activeBatches = new();

    /// <summary>
    /// Open a session for a workbook (session lifecycle start)
    /// </summary>
    /// <param name="args">Command arguments: open file.xlsx</param>
    /// <returns>0 for success, 1 for error</returns>
    public int Open(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] open <file.xlsx>");
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

            AnsiConsole.MarkupLine($"[green]✓[/] [bold]Session opened[/]");
            AnsiConsole.MarkupLine($"[cyan]Session ID:[/] {batchId}");
            AnsiConsole.MarkupLine($"[dim]File:[/] {normalizedPath}");
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[bold]Next steps:[/]");
            AnsiConsole.MarkupLine($"[dim]• Use --session-id {batchId} with any command[/]");
            AnsiConsole.MarkupLine($"[dim]• All operations use same Excel instance[/]");
            AnsiConsole.MarkupLine($"[dim]• Call save {batchId} when done[/]");

            return 0;
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }

    /// <summary>
    /// Save the workbook (keeps session open)
    /// </summary>
    /// <param name="args">Command arguments: save session-id</param>
    /// <returns>0 for success, 1 for error</returns>
    public int Save(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] save <session-id>");
            return 1;
        }

        string sessionId = args[1];

        try
        {
            // Retrieve batch session (don't remove it - keep session open)
            if (!_activeBatches.TryGetValue(sessionId, out var batch))
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Session '{sessionId}' not found");
                AnsiConsole.MarkupLine("[yellow]Hint:[/] Use list to see active sessions");
                return 1;
            }

            string filePath = batch.WorkbookPath;

            try
            {
                var task = Task.Run(async () => await batch.SaveAsync());
                task.GetAwaiter().GetResult();

                AnsiConsole.MarkupLine($"[green]✓[/] [bold]Changes saved[/]");
                AnsiConsole.MarkupLine($"[cyan]Session ID:[/] {sessionId}");
                AnsiConsole.MarkupLine($"[dim]File:[/] {filePath}");
                AnsiConsole.MarkupLine("[dim]Session remains open for more operations[/]");

                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error saving:[/] {ex.Message.EscapeMarkup()}");
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
    /// Close a session (session lifecycle end, discards changes)
    /// </summary>
    /// <param name="args">Command arguments: close session-id</param>
    /// <returns>0 for success, 1 for error</returns>
    public int Close(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] close <session-id>");
            return 1;
        }

        string sessionId = args[1];

        try
        {
            // Retrieve batch session and remove it
            if (!_activeBatches.TryRemove(sessionId, out var batch))
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Session '{sessionId}' not found");
                AnsiConsole.MarkupLine("[yellow]Hint:[/] Use list to see active sessions");
                return 1;
            }

            string filePath = batch.WorkbookPath;

            try
            {
                var task = Task.Run(async () => await batch.DisposeAsync());
                task.GetAwaiter().GetResult();

                AnsiConsole.MarkupLine($"[green]✓[/] [bold]Session closed[/]");
                AnsiConsole.MarkupLine($"[cyan]Session ID:[/] {sessionId}");
                AnsiConsole.MarkupLine($"[dim]File:[/] {filePath}");
                AnsiConsole.MarkupLine("[yellow]Changes discarded[/]");

                return 0;
            }
            catch
            {
                // If dispose fails, try to dispose anyway to prevent resource leaks
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
    /// List all active sessions
    /// </summary>
    /// <param name="args">Command arguments: list</param>
    /// <returns>0 for success</returns>
#pragma warning disable IDE0060 // Remove unused parameter
    public int List(string[] args)
#pragma warning restore IDE0060 // Remove unused parameter
    {
        var sessions = _activeBatches.ToList();

        if (sessions.Count == 0)
        {
            AnsiConsole.MarkupLine("[yellow]No active sessions[/]");
            AnsiConsole.MarkupLine("[dim]Start one with:[/] [cyan]open file.xlsx[/]");
            return 0;
        }

        AnsiConsole.MarkupLine($"[bold]Active Sessions:[/] {sessions.Count}\n");

        var table = new Table();
        table.AddColumn("[bold]Session ID[/]");
        table.AddColumn("[bold]File Path[/]");

        foreach (var kvp in sessions)
        {
            table.AddRow(
                $"[cyan]{kvp.Key}[/]",
                $"[dim]{kvp.Value.WorkbookPath.EscapeMarkup()}[/]"
            );
        }

        AnsiConsole.Write(table);
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[yellow]⚠[/] [bold]Remember to close sessions![/]");
        AnsiConsole.MarkupLine("[dim]Each session holds Excel open. Call close to release resources.[/]");

        return 0;
    }

    /// <summary>
    /// Get an active session by ID.
    /// Used internally by other commands to retrieve the session for operations.
    /// </summary>
    internal static IExcelBatch? GetSession(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return null;
        }

        return _activeBatches.TryGetValue(sessionId, out var session) ? session : null;
    }

    /// <summary>
    /// Legacy method for backward compatibility - forwards to GetSession
    /// </summary>
    internal static IExcelBatch? GetBatch(string batchId) => GetSession(batchId);
}
