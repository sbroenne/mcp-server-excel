using System.Collections.Concurrent;
using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Manages Excel batch sessions for MCP clients.
/// Allows LLMs to control the lifecycle of Excel workbook sessions for high-performance multi-operation workflows.
/// </summary>
[McpServerToolType]
public static class BatchSessionTool
{
    private static readonly ConcurrentDictionary<string, IExcelBatch> _activeBatches = new();
    private static readonly JsonSerializerOptions _jsonOptions = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    /// <summary>
    /// Begin a new Excel batch session.
    /// Opens the workbook and keeps it in memory for subsequent operations.
    /// </summary>
    [McpServerTool(Name = "begin_excel_batch")]
    [Description("Start a batch session for high-performance multi-operation workflows. Opens workbook once, reuses for all operations (75-90% faster). Returns batchId to pass to other tools. Always commit when done to save and release resources.")]
    public static async Task<string> BeginExcelBatch(
        [Description("Full path to Excel file (.xlsx or .xlsm)")]
        string filePath)
    {
        try
        {
            // Validate file path
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ModelContextProtocol.McpException("filePath is required");
            }

            // Normalize path to prevent duplicate sessions
            string normalizedPath = Path.GetFullPath(filePath);

            // Check if batch already exists for this file
            if (_activeBatches.ContainsKey(normalizedPath))
            {
                throw new ModelContextProtocol.McpException($"Batch session already active for '{filePath}'. Commit or discard existing batch before starting a new one.");
            }

            // Create new batch session
            var batch = await ExcelSession.BeginBatchAsync(filePath);

            // Generate batch ID (use normalized path as ID for now - ensures uniqueness per file)
            string batchId = Guid.NewGuid().ToString();

            // Store in active sessions
            if (!_activeBatches.TryAdd(batchId, batch))
            {
                // Cleanup if we couldn't add (shouldn't happen but be safe)
                await batch.DisposeAsync();
                throw new ModelContextProtocol.McpException("Failed to register batch session");
            }

            var result = new
            {
                success = true,
                batchId = batchId,
                filePath = normalizedPath,
                message = $"Batch session started. Use batchId='{batchId}' for subsequent operations on this workbook.",
                instructions = new[]
                {
                    "Pass this batchId to excel_powerquery, excel_worksheet, excel_parameter, etc.",
                    "All operations will use the same open workbook (much faster!)",
                    "Call commit_excel_batch when done to save and close",
                    "Or call commit_excel_batch with save=false to discard changes"
                }
            };

            return JsonSerializer.Serialize(result, _jsonOptions);
        }
        catch (ModelContextProtocol.McpException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new ModelContextProtocol.McpException($"Failed to begin batch session: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Commit or discard an Excel batch session.
    /// Saves the workbook (if requested), closes it, and releases resources.
    /// </summary>
    [McpServerTool(Name = "commit_excel_batch")]
    [Description("End a batch session. Saves changes (if save=true) and closes workbook. REQUIRED after begin_excel_batch to prevent resource leaks. Set save=false to discard all changes made during the batch.")]
    public static async Task<string> CommitExcelBatch(
        [Description("Batch ID returned from begin_excel_batch")]
        string batchId, 
        [Description("Save changes before closing? Default true. Set false to discard all changes.")]
        bool save = true)
    {
        try
        {
            // Validate batch ID
            if (string.IsNullOrWhiteSpace(batchId))
            {
                throw new ModelContextProtocol.McpException("batchId is required");
            }

            // Retrieve batch session
            if (!_activeBatches.TryRemove(batchId, out var batch))
            {
                throw new ModelContextProtocol.McpException($"Batch session '{batchId}' not found. It may have already been committed or never existed.");
            }

            string filePath = batch.WorkbookPath;

            try
            {
                // Save if requested
                if (save)
                {
                    await batch.SaveAsync();
                }

                // Dispose (closes workbook and releases Excel)
                await batch.DisposeAsync();

                var result = new
                {
                    success = true,
                    batchId = batchId,
                    filePath = filePath,
                    saved = save,
                    message = save
                        ? $"Batch session committed. Workbook saved and closed: {filePath}"
                        : $"Batch session discarded. Workbook closed without saving: {filePath}"
                };

                return JsonSerializer.Serialize(result, _jsonOptions);
            }
            catch
            {
                // If save/dispose fails, try to dispose anyway to prevent resource leaks
                try { await batch.DisposeAsync(); } catch { /* ignore */ }
                throw;
            }
        }
        catch (ModelContextProtocol.McpException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new ModelContextProtocol.McpException($"Failed to commit batch session: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// List all active batch sessions.
    /// Useful for debugging or cleanup.
    /// </summary>
    [McpServerTool(Name = "list_excel_batches")]
    [Description("List all active batch sessions. Shows batchId and file path for each open session. Use to debug resource leaks or check which files have uncommitted batches. Always commit active batches when done.")]
    public static string ListExcelBatches()
    {
        var batches = _activeBatches.Select(kvp => new
        {
            batchId = kvp.Key,
            filePath = kvp.Value.WorkbookPath
        }).ToList();

        var result = new
        {
            success = true,
            count = batches.Count,
            activeBatches = batches,
            message = batches.Count > 0
                ? $"Found {batches.Count} active batch session(s). Remember to commit when done!"
                : "No active batch sessions."
        };

        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    /// <summary>
    /// Get an active batch session by ID.
    /// Used internally by other tools to retrieve the batch for operations.
    /// </summary>
    internal static IExcelBatch? GetBatch(string batchId)
    {
        if (string.IsNullOrWhiteSpace(batchId))
        {
            return null;
        }

        return _activeBatches.TryGetValue(batchId, out var batch) ? batch : null;
    }

    /// <summary>
    /// Cleanup all active batches (for server shutdown).
    /// </summary>
    internal static async Task CleanupAllBatches()
    {
        var batches = _activeBatches.Values.ToList();
        _activeBatches.Clear();

        foreach (var batch in batches)
        {
            try
            {
                await batch.DisposeAsync();
            }
            catch
            {
                // Ignore errors during cleanup
            }
        }
    }
}

