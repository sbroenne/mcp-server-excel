using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Helper methods for CLI commands to support session mode
/// </summary>
internal static class CommandHelper
{
    /// <summary>
    /// Executes an async Core command with optional session management.
    /// If --session-id is provided in args, uses existing session. Otherwise, creates session-of-one.
    /// </summary>
    /// <typeparam name="T">Return type of the command</typeparam>
    /// <param name="args">Command arguments (may contain --session-id parameter)</param>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="save">Whether to save changes (only used for session-of-one)</param>
    /// <param name="action">Async action that takes IExcelBatch and returns Task&lt;T&gt;</param>
    /// <returns>Result of the command</returns>
    public static T WithBatchAsync<T>(
        string[] args,
        string filePath,
        bool save,
        Func<IExcelBatch, Task<T>> action)
    {
        // Check for --session-id parameter
        string? sessionId = GetSessionIdFromArgs(args);

        if (!string.IsNullOrEmpty(sessionId))
        {
            // Use existing session
            var batch = BatchCommands.GetBatch(sessionId);
            if (batch == null)
            {
                throw new InvalidOperationException(
                    $"Session '{sessionId}' not found. Use list to see active sessions.");
            }

            // Verify file path matches session
            if (!string.Equals(batch.WorkbookPath, Path.GetFullPath(filePath), StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"File path mismatch. Session is for '{batch.WorkbookPath}' but operation requested '{filePath}'.");
            }

            // Execute in existing session (no save here, that's done in save command)
            var task = Task.Run(async () => await action(batch));
            return task.GetAwaiter().GetResult();
        }
        else
        {
            // Session-of-one (current behavior)
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
    /// Extracts the session ID from command arguments
    /// </summary>
    /// <param name="args">Command arguments</param>
    /// <returns>Session ID if found, null otherwise</returns>
    private static string? GetSessionIdFromArgs(string[] args)
    {
        for (int i = 0; i < args.Length - 1; i++)
        {
            if (args[i].Equals("--session-id", StringComparison.OrdinalIgnoreCase))
            {
                return args[i + 1];
            }
        }
        return null;
    }
}
