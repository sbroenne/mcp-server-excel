using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - Workflow hints are contextual per-call

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel named range (parameter) operations.
/// </summary>
[McpServerToolType]
public static class ExcelNamedRangeTool
{
    // Cache JsonSerializerOptions to satisfy CA1869
    private static readonly JsonSerializerOptions s_jsonOptions = new() { PropertyNameCaseInsensitive = true };

    // Cache suggestedNextActions arrays to satisfy CA1861
    private static readonly string[] s_getNextActions =
    [
        "Use 'set' to update this parameter value",
        "Use this value in excel_range or excel_powerquery operations",
        "Use 'update' to change the cell reference"
    ];

    private static readonly string[] s_createBulkNextActions =
    [
        "Use 'list' to verify all created named ranges",
        "Use 'set' to assign initial values",
        "Use excel_range to populate data in named range regions"
    ];

    /// <summary>
    /// Manage Excel parameters (named ranges) - configuration values and reusable references
    /// </summary>
    [McpServerTool(Name = "excel_namedrange")]
    [Description(@"Manage Excel named ranges")]
    public static async Task<string> ExcelParameter(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        NamedRangeAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [Description("Session ID from excel_file 'open' action")]
        string sessionId,

        [StringLength(255, MinimumLength = 1)]
        [Description("Named range name (for get, set, create, update, delete actions)")]
        string? namedRangeName = null,

        [Description("Named range value (for set action) or cell reference (for create/update actions, e.g., 'Sheet1!A1')")]
        string? value = null,

        [Description("JSON array of named ranges for create-bulk action: [{name: 'Name', reference: 'Sheet1!A1', value: 'text'}, ...]")]
        string? namedRangesJson = null)
    {
        try
        {
            var namedRangeCommands = new NamedRangeCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                NamedRangeAction.List => await ListNamedRangesAsync(namedRangeCommands, sessionId),
                NamedRangeAction.Get => await GetNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName),
                NamedRangeAction.Set => await SetNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName, value),
                NamedRangeAction.Create => await CreateNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName, value),
                NamedRangeAction.CreateBulk => await CreateBulkNamedRangesAsync(namedRangeCommands, sessionId, namedRangesJson),
                NamedRangeAction.Update => await UpdateNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName, value),
                NamedRangeAction.Delete => await DeleteNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName),
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

    private static async Task<string> ListNamedRangesAsync(NamedRangeCommands commands, string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ListAsync(batch));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        var count = result.NamedRanges?.Count ?? 0;

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
                    "Add more operations in this session"
                }
                :
                [
                    "Use 'get' to retrieve named range values",
                    "Use 'set' to update parameter values",
                    "Use excel_range with sheetName='' to reference named ranges in formulas",
                    "Continue operations in this session"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName)
    {
        if (string.IsNullOrEmpty(namedRangeName))
            throw new ModelContextProtocol.McpException("namedRangeName is required for get action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetAsync(batch, namedRangeName));

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

    private static async Task<string> SetNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName, string? value)
    {
        if (string.IsNullOrEmpty(namedRangeName) || value == null)
            throw new ModelContextProtocol.McpException("namedRangeName and value are required for set action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SetAsync(batch, namedRangeName, value));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        return JsonSerializer.Serialize(new
        {
            result.Success,
            workflowHint = $"Parameter '{namedRangeName}' updated to '{value}'.",
            suggestedNextActions = new[]
            {
                "Use 'get' to verify the new value",
                "Use excel_range to see how formulas using this parameter changed",
                "Set more parameters in this session"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName, string? value)
    {
        if (string.IsNullOrEmpty(namedRangeName) || string.IsNullOrEmpty(value))
            throw new ModelContextProtocol.McpException("namedRangeName and value (cell reference) are required for update action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.UpdateAsync(batch, namedRangeName, value));

        // Always return JSON (success or failure) - MCP clients handle the success flag
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
                    "Update more references in this session"
                }
                :
                [
                    "Use 'list' to verify named range exists",
                    "Check cell reference format (e.g., 'Sheet1!A1' or 'Sheet1!A1:B10')",
                    "Ensure target sheet exists in workbook"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName, string? value)
    {
        if (string.IsNullOrEmpty(namedRangeName) || string.IsNullOrEmpty(value))
            throw new ModelContextProtocol.McpException("namedRangeName and value (cell reference) are required for create action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.CreateAsync(batch, namedRangeName, value));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        return JsonSerializer.Serialize(new
        {
            result.Success,
            workflowHint = $"Named range '{namedRangeName}' created pointing to {value}.",
            suggestedNextActions = new[]
            {
                "Use 'set' to assign an initial value",
                "Use 'get' to verify the named range",
                "Create more named ranges in this session"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName)
    {
        if (string.IsNullOrEmpty(namedRangeName))
            throw new ModelContextProtocol.McpException("namedRangeName is required for delete action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.DeleteAsync(batch, namedRangeName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
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
                    "Delete more named ranges in this session"
                }
                :
                [
                    "Use 'list' to verify named range exists",
                    "Check if workbook or sheet is protected",
                    "Verify named range name spelling is correct"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateBulkNamedRangesAsync(NamedRangeCommands commands, string sessionId, string? namedRangesJson)
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.CreateBulkAsync(batch, parameters));

        // Add workflow hints (CreateBulk returns OperationResult, not specialized type)
        return JsonSerializer.Serialize(new
        {
            result.Success,
            workflowHint = "Bulk named range creation completed.",
            suggestedNextActions = s_createBulkNextActions
        }, ExcelToolsBase.JsonOptions);
    }
}
