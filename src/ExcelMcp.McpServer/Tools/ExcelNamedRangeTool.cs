using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - Workflow hints are contextual per-call
#pragma warning disable IDE0060 // batchId parameter kept for compatibility, will be removed in final cleanup phase

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel named range (parameter) operations.
/// </summary>
[McpServerToolType]
public static class ExcelNamedRangeTool
{
    // Cache suggestedNextActions arrays to satisfy CA1861
    private static readonly string[] s_getNextActions =
    [
        "Use 'set' to update this parameter value",
        "Use this value in excel_range or excel_powerquery operations",
        "Use 'update' to change the cell reference"
    ];

    /// <summary>
    /// Manage Excel parameters (named ranges) - configuration values and reusable references
    /// </summary>
    [McpServerTool(Name = "excel_namedrange")]
    [Description(@"Manage Excel named ranges as parameters (configuration values).

USE CASES:
- Configuration values: StartDate, ReportYear, Threshold
- Reusable parameters: Formula inputs, dynamic ranges
- Power Query parameters: Reference in M code via Excel.CurrentWorkbook()

⚡ PERFORMANCE: File handle caching automatically optimizes sequential operations.

⭐ BULK OPERATIONS: Use 'create-bulk' action for even better efficiency (one call for multiple parameters).

RELATED TOOLS:
- excel_range: For bulk data operations on named range contents
- excel_powerquery: To reference named ranges in M code

Optional batchId for batch sessions.")]
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
        var result = await commands.ListAsync(filePath);

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
                    inBatch ? "Create additional named ranges as needed" : "File handle caching provides automatic performance optimization"
                }
                :
                [
                    "Use 'get' to retrieve named range values",
                    "Use 'set' to update parameter values",
                    "Use excel_range with sheetName='' to reference named ranges in formulas",
                    inBatch ? "Continue with additional operations" : count > 3 ? "File handle caching provides automatic performance optimization" : "Use 'update' to change cell references"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetNamedRangeAsync(NamedRangeCommands commands, string filePath, string? namedRangeName, string? batchId)
    {
        if (string.IsNullOrEmpty(namedRangeName))
            throw new ModelContextProtocol.McpException("namedRangeName is required for get action");

        var result = await commands.GetAsync(filePath, namedRangeName);

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

        var result = await commands.SetAsync(filePath, namedRangeName, value);
        await ComInterop.Session.FileHandleManager.Instance.SaveAsync(filePath);

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
                inBatch ? "Continue with additional operations" : "Set more parameters? Use excel_batch (90% faster)"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateNamedRangeAsync(NamedRangeCommands commands, string filePath, string? namedRangeName, string? value, string? batchId)
    {
        if (string.IsNullOrEmpty(namedRangeName) || string.IsNullOrEmpty(value))
            throw new ModelContextProtocol.McpException("namedRangeName and value (cell reference) are required for update action");

        var result = await commands.UpdateAsync(filePath, namedRangeName, value);
        await ComInterop.Session.FileHandleManager.Instance.SaveAsync(filePath);

        // Always return JSON (success or failure) - MCP clients handle the success flag
        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Named range '{namedRangeName}' now points to {value}. Formulas referencing it will use new location."
                : $"Failed to update '{namedRangeName}'. Verify the named range exists and cell reference is valid.",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Use 'get' to retrieve value from new location",
                    "Use excel_range to read data from new cell reference",
                    inBatch ? "Continue with additional operations" : "Update more references? Use excel_batch for efficiency"
                }
                :
                [
                    "Use 'list' to verify named range exists",
                    "Check cell reference format (e.g., 'Sheet1!A1' or 'Sheet1!A1:B10')",
                    "Ensure target sheet exists in workbook"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateNamedRangeAsync(NamedRangeCommands commands, string filePath, string? namedRangeName, string? value, string? batchId)
    {
        if (string.IsNullOrEmpty(namedRangeName) || string.IsNullOrEmpty(value))
            throw new ModelContextProtocol.McpException("namedRangeName and value (cell reference) are required for create action");

        var result = await commands.CreateAsync(filePath, namedRangeName, value);
        await ComInterop.Session.FileHandleManager.Instance.SaveAsync(filePath);

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

        var result = await commands.DeleteAsync(filePath, namedRangeName);
        await ComInterop.Session.FileHandleManager.Instance.SaveAsync(filePath);

        // Always return JSON (success or failure) - MCP clients handle the success flag
        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Named range '{namedRangeName}' deleted successfully. Formulas referencing it will show #NAME? error."
                : $"Failed to delete '{namedRangeName}'. Verify the named range exists and is not protected.",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Use 'list' to verify deletion",
                    "Check formulas that referenced this parameter (will show #NAME? errors)",
                    inBatch ? "Continue with additional operations" : "Delete more named ranges? Use excel_batch for efficiency"
                }
                :
                [
                    "Use 'list' to verify named range exists",
                    "Check if workbook or sheet is protected",
                    "Verify named range name spelling is correct"
                ]
        }, ExcelToolsBase.JsonOptions);
    }
}
