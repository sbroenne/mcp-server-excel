using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Helper methods for CLI commands to support batch mode
/// </summary>
internal static class CommandHelper
{
    /// <summary>
    /// Executes an async Core command with optional batch session management.
    /// If --batch-id is provided in args, uses existing batch session. Otherwise, creates batch-of-one.
    /// </summary>
    /// <typeparam name="T">Return type of the command</typeparam>
    /// <param name="args">Command arguments (may contain --batch-id parameter)</param>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="save">Whether to save changes (only used for batch-of-one)</param>
    /// <param name="action">Async action that takes IExcelBatch and returns Task&lt;T&gt;</param>
    /// <returns>Result of the command</returns>
    public static T WithBatchAsync<T>(
        string[] args,
        string filePath,
        bool save,
        Func<IExcelBatch, Task<T>> action)
    {
        // Check for --batch-id parameter
        string? batchId = GetBatchIdFromArgs(args);

        if (!string.IsNullOrEmpty(batchId))
        {
            // Use existing batch session
            var batch = BatchCommands.GetBatch(batchId);
            if (batch == null)
            {
                throw new InvalidOperationException(
                    $"Batch session '{batchId}' not found. Use batch-list to see active sessions.");
            }

            // Verify file path matches batch
            if (!string.Equals(batch.WorkbookPath, Path.GetFullPath(filePath), StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"File path mismatch. Batch session is for '{batch.WorkbookPath}' but operation requested '{filePath}'.");
            }

            // Execute in existing batch (no save here, that's done in batch-commit)
            var task = Task.Run(async () => await action(batch));
            return task.GetAwaiter().GetResult();
        }
        else
        {
            // Batch-of-one (current behavior)
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var result = await action(batch);

                if (save)
                {
                    await batch.SaveAsync();
                }

                return result;
            });
            return task.GetAwaiter().GetResult();
        }
    }

    /// <summary>
    /// Extracts the batch ID from command arguments
    /// </summary>
    /// <param name="args">Command arguments</param>
    /// <returns>Batch ID if found, null otherwise</returns>
    private static string? GetBatchIdFromArgs(string[] args)
    {
        for (int i = 0; i < args.Length - 1; i++)
        {
            if (args[i].Equals("--batch-id", StringComparison.OrdinalIgnoreCase))
            {
                return args[i + 1];
            }
        }
        return null;
    }
}
