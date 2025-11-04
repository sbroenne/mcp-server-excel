using System.Collections.Concurrent;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel batch session management tool for MCP server.
/// Allows LLMs to control the lifecycle of Excel workbook sessions for high-performance multi-operation workflows.
/// </summary>
[McpServerToolType]
public static class BatchSessionTool
{
    private static readonly ConcurrentDictionary<string, IExcelBatch> _activeBatches = new();
    private static readonly JsonSerializerOptions _jsonOptions = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    /// <summary>
    /// Manage Excel batch sessions for multi-operation workflows
    /// </summary>
    [McpServerTool(Name = "excel_batch")]
    [Description(@"Manage batch sessions for high-performance workflows (75-90% faster).

Actions: begin, commit, list

Use begin to start a session, commit to end it, list to debug.")]
    public static async Task<string> ExcelBatch(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        BatchAction action,

        [Description("Full path to Excel file (.xlsx or .xlsm) - required for 'begin' action")]
        string? filePath = null,

        [Description("Batch ID from begin action - required for 'commit' action")]
        string? batchId = null,

        [Description("Save changes before closing? Default true. Set false to discard - used with 'commit' action")]
        bool save = true,
        
        [Description("Timeout in minutes for batch operations. Default: 2 minutes for save operations")]
        double? timeout = null)
    {
        try
        {
            return action switch
            {
                BatchAction.Begin => await BeginBatchAsync(filePath!),
                BatchAction.Commit => await CommitBatchAsync(batchId!, save, timeout),
                BatchAction.List => ListBatches(),
                _ => throw new ModelContextProtocol.McpException($"Unknown action: {action} ({action.ToActionString()})")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw;
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), filePath ?? batchId ?? "");
            throw; // Unreachable but satisfies compiler
        }
    }

    private static async Task<string> BeginBatchAsync(string filePath)
    {
        // Validate file path
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ModelContextProtocol.McpException("filePath is required for 'begin' action");
        }

        // Normalize path to prevent duplicate sessions
        string normalizedPath = Path.GetFullPath(filePath);

        // Check if file exists BEFORE attempting to open
        if (!File.Exists(normalizedPath))
        {
            throw new ModelContextProtocol.McpException(
                $"File not found: '{normalizedPath}'. " +
                $"Use excel_file(action: 'create-empty', filePath: '{normalizedPath}') to create it first, " +
                $"or verify the path is correct.");
        }

        // Check if batch already exists for this file
        if (_activeBatches.ContainsKey(normalizedPath))
        {
            throw new ModelContextProtocol.McpException($"Batch session already active for '{filePath}'. Commit or discard existing batch before starting a new one.");
        }

        // Create new batch session
        var batch = await ExcelSession.BeginBatchAsync(filePath);

        // Generate batch ID
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
            message = $"Batch session started. Use batchId='{batchId}' for subsequent operations.",
            suggestedNextActions = new[]
            {
                "Pass batchId to excel_powerquery, excel_worksheet, excel_namedrange, etc.",
                "All operations use same Excel instance (75-90% faster!)",
                "Call excel_batch(action: 'commit', batchId: '...') when done"
            },
            workflowHint = "Batch active. Use batchId for all operations, commit when done."
        };

        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    private static async Task<string> CommitBatchAsync(string batchId, bool save, double? timeoutMinutes)
    {
        // Validate batch ID
        if (string.IsNullOrWhiteSpace(batchId))
        {
            throw new ModelContextProtocol.McpException("batchId is required for 'commit' action");
        }

        // Retrieve batch session
        if (!_activeBatches.TryRemove(batchId, out var batch))
        {
            throw new ModelContextProtocol.McpException($"Batch session '{batchId}' not found. It may have already been committed or never existed.");
        }

        string filePath = batch.WorkbookPath;

        try
        {
            // Save if requested (with extended timeout for large workbooks)
            if (save)
            {
                var timeoutSpan = timeoutMinutes.HasValue ? (TimeSpan?)TimeSpan.FromMinutes(timeoutMinutes.Value) : null;
                await batch.SaveAsync(timeout: timeoutSpan);
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
                    ? $"Batch committed. Workbook saved: {filePath}"
                    : $"Batch discarded. Workbook closed without saving: {filePath}",
                workflowHint = save
                    ? "Changes saved. Batch complete."
                    : "Changes discarded. Batch complete."
            };

            return JsonSerializer.Serialize(result, _jsonOptions);
        }
        catch (TimeoutException ex)
        {
            // Timeout during save - provide LLM guidance
            var result = new
            {
                success = false,
                errorMessage = ex.Message,
                filePath = filePath,
                action = "commit",
                suggestedNextActions = new[]
                {
                    "Large workbook detected - save operation timed out",
                    "Try again - timeout was likely transient",
                    "Check if Excel is showing a dialog or prompt",
                    "Verify the file is not locked by another process",
                    "Consider using save=false to discard changes if stuck"
                },
                operationContext = new Dictionary<string, object>
                {
                    { "OperationType", "BatchSession.Commit" },
                    { "SaveRequested", save },
                    { "TimeoutReached", true },
                    { "WorkbookPath", filePath }
                },
                isRetryable = !ex.Message.Contains("maximum timeout"),
                retryGuidance = ex.Message.Contains("maximum timeout")
                    ? "Maximum timeout reached. Check for Excel dialogs or file locks manually."
                    : "Retry acceptable - timeout may have been transient."
            };

            return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
        }
        catch
        {
            // If save/dispose fails, try to dispose anyway to prevent resource leaks
            try { await batch.DisposeAsync(); } catch { /* ignore */ }
            throw;
        }
    }

    private static string ListBatches()
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
                ? $"Found {batches.Count} active batch session(s). Remember to commit!"
                : "No active batches.",
            suggestedNextActions = batches.Count > 0
                ? new[] { "Commit batches with excel_batch(action: 'commit', batchId: '...')" }
                : new[] { "Start batch with excel_batch(action: 'begin', filePath: '...')" },
            workflowHint = batches.Count > 0
                ? "Active batches found. Always commit to prevent resource leaks."
                : "No active batches."
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

