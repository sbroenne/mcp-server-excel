using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel parameter (named range) management tool for MCP server.
/// Handles named ranges as configuration parameters for Excel automation.
///
/// LLM Usage Patterns:
/// - Use "list" to see all named ranges (parameters) in a workbook
/// - Use "get" to retrieve Named range values for configuration
/// - Use "set" to update Named range values for dynamic behavior
/// - Use "update" to change parameter cell reference
/// - Use "create" to define new named ranges as parameters
/// - Use "delete" to remove obsolete parameters
///
/// Note: Named ranges are Excel's way of creating reusable parameters that can be
/// referenced in formulas and Power Query. They're ideal for configuration values.
/// </summary>
[McpServerToolType]
public static class ExcelNamedRangeTool
{
    // Cache JsonSerializerOptions to satisfy CA1869
    private static readonly JsonSerializerOptions s_jsonOptions = new() { PropertyNameCaseInsensitive = true };

    // Cache suggestedNextActions arrays to satisfy CA1861
    private static readonly string[] s_getNextActions = new[]
    {
        "Use 'set' to update this parameter value",
        "Use this value in excel_range or excel_powerquery operations",
        "Use 'update' to change the cell reference"
    };

    private static readonly string[] s_createBulkNextActions = new[]
    {
        "Use 'list' to verify all created named ranges",
        "Use 'set' to assign initial values",
        "Use excel_range to populate data in named range regions"
    };

    /// <summary>
    /// Manage Excel parameters (named ranges) - configuration values and reusable references
    /// </summary>
    [McpServerTool(Name = "excel_namedrange")]
    [Description(@"Manage Excel named ranges as parameters (configuration values).

⚡ PERFORMANCE: For creating 2+ parameters, use begin_excel_batch FIRST (90% faster):
  1. batch = begin_excel_batch(excelPath: 'file.xlsx')
  2. excel_namedrange(action: 'create', ..., batchId: batch.batchId)  // repeat for each parameter
  3. commit_excel_batch(batchId: batch.batchId, save: true)

⭐ NEW: Use 'create-bulk' action for even better efficiency (one call for multiple parameters).

Actions available as dropdown in MCP clients.")]
    public static async Task<string> ExcelParameter(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        NamedRangeAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [StringLength(255, MinimumLength = 1)]
        [Description("Named range name (for get, set, create, update, delete actions)")]
        string? namedRangeName = null,

        [Description("Named range value (for set action) or cell reference (for create/update actions, e.g., 'Sheet1!A1')")]
        string? value = null,

        [Description("JSON array of named ranges for create-bulk action: [{name: 'Name', reference: 'Sheet1!A1', value: 'text'}, ...]")]
        string? namedRangesJson = null,

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var namedRangeCommands = new NamedRangeCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                NamedRangeAction.List => await ListNamedRangesAsync(namedRangeCommands, excelPath, batchId),
                NamedRangeAction.Get => await GetNamedRangeAsync(namedRangeCommands, excelPath, namedRangeName, batchId),
                NamedRangeAction.Set => await SetNamedRangeAsync(namedRangeCommands, excelPath, namedRangeName, value, batchId),
                NamedRangeAction.Create => await CreateNamedRangeAsync(namedRangeCommands, excelPath, namedRangeName, value, batchId),
                NamedRangeAction.CreateBulk => await CreateBulkNamedRangesAsync(namedRangeCommands, excelPath, namedRangesJson, batchId),
                NamedRangeAction.Update => await UpdateNamedRangeAsync(namedRangeCommands, excelPath, namedRangeName, value, batchId),
                NamedRangeAction.Delete => await DeleteNamedRangeAsync(namedRangeCommands, excelPath, namedRangeName, batchId),
                _ => throw new ModelContextProtocol.McpException($"Unknown action: {action} ({action.ToActionString()})")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);
            throw; // Unreachable but satisfies compiler
        }
    }

    private static async Task<string> ListNamedRangesAsync(NamedRangeCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListAsync(batch));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        var count = result.NamedRanges?.Count ?? 0;
        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.NamedRanges,
            workflowHint = count == 0
                ? "No named ranges found. Create parameters for reusable values and formula references."
                : $"Found {count} named range(s). Use 'get' to retrieve values or 'set' to update them.",
            suggestedNextActions = count == 0
                ? new[]
                {
                    "Use 'create' to define new named ranges as parameters",
                    inBatch ? "Add more operations in this batch session" : "Use excel_batch for creating multiple parameters (90% faster)"
                }
                : new[]
                {
                    "Use 'get' to retrieve named range values",
                    "Use 'set' to update parameter values",
                    "Use excel_range with sheetName='' to reference named ranges in formulas",
                    inBatch ? "Continue batch operations" : count > 3 ? "Use excel_batch for bulk updates (90% faster)" : "Use 'update' to change cell references"
                }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetNamedRangeAsync(NamedRangeCommands commands, string filePath, string? namedRangeName, string? batchId)
    {
        if (string.IsNullOrEmpty(namedRangeName))
            throw new ModelContextProtocol.McpException("namedRangeName is required for get action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetAsync(batch, namedRangeName));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.NamedRangeName,
            result.Value,
            workflowHint = $"Retrieved value: {result.Value}",
            suggestedNextActions = s_getNextActions
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetNamedRangeAsync(NamedRangeCommands commands, string filePath, string? namedRangeName, string? value, string? batchId)
    {
        if (string.IsNullOrEmpty(namedRangeName) || value == null)
            throw new ModelContextProtocol.McpException("namedRangeName and value are required for set action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.SetAsync(batch, namedRangeName, value));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            workflowHint = $"Parameter '{namedRangeName}' updated to '{value}'.",
            suggestedNextActions = new[]
            {
                "Use 'get' to verify the new value",
                "Use excel_range to see how formulas using this parameter changed",
                inBatch ? "Continue batch operations" : "Set more parameters? Use excel_batch (90% faster)"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateNamedRangeAsync(NamedRangeCommands commands, string filePath, string? namedRangeName, string? value, string? batchId)
    {
        if (string.IsNullOrEmpty(namedRangeName) || string.IsNullOrEmpty(value))
            throw new ModelContextProtocol.McpException("namedRangeName and value (cell reference) are required for update action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.UpdateAsync(batch, namedRangeName, value));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateNamedRangeAsync(NamedRangeCommands commands, string filePath, string? namedRangeName, string? value, string? batchId)
    {
        if (string.IsNullOrEmpty(namedRangeName) || string.IsNullOrEmpty(value))
            throw new ModelContextProtocol.McpException("namedRangeName and value (cell reference) are required for create action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.CreateAsync(batch, namedRangeName, value));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            workflowHint = $"Named range '{namedRangeName}' created pointing to {value}.",
            suggestedNextActions = new[]
            {
                "Use 'set' to assign an initial value",
                "Use 'get' to verify the named range",
                inBatch ? "Create more named ranges in this batch" : "Creating multiple? Use excel_batch or 'create-bulk' (90% faster)"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteNamedRangeAsync(NamedRangeCommands commands, string filePath, string? namedRangeName, string? batchId)
    {
        if (string.IsNullOrEmpty(namedRangeName))
            throw new ModelContextProtocol.McpException("namedRangeName is required for delete action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.DeleteAsync(batch, namedRangeName));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateBulkNamedRangesAsync(NamedRangeCommands commands, string excelPath, string? namedRangesJson, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(namedRangesJson))
            throw new ModelContextProtocol.McpException("namedRangesJson is required for create-bulk action");

        // Deserialize JSON array of named range definitions
        List<NamedRangeDefinition>? parameters;
        try
        {
            parameters = JsonSerializer.Deserialize<List<NamedRangeDefinition>>(
                namedRangesJson,
                s_jsonOptions);

            if (parameters == null || parameters.Count == 0)
                throw new ModelContextProtocol.McpException("namedRangesJson must contain at least one named range definition");
        }
        catch (JsonException ex)
        {
            throw new ModelContextProtocol.McpException($"Invalid namedRangesJson format: {ex.Message}");
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.CreateBulkAsync(batch, parameters));

        if (!result.Success)
        {
            throw new ModelContextProtocol.McpException($"create-bulk failed: {result.ErrorMessage}");
        }

        // Add workflow hints (CreateBulk returns OperationResult, not specialized type)
        return JsonSerializer.Serialize(new
        {
            result.Success,
            workflowHint = "Bulk named range creation completed.",
            suggestedNextActions = s_createBulkNextActions
        }, ExcelToolsBase.JsonOptions);
    }
}
