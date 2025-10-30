using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel parameter (named range) management tool for MCP server.
/// Handles named ranges as configuration parameters for Excel automation.
///
/// LLM Usage Patterns:
/// - Use "list" to see all named ranges (parameters) in a workbook
/// - Use "get" to retrieve parameter values for configuration
/// - Use "set" to update parameter values for dynamic behavior
/// - Use "update" to change parameter cell reference
/// - Use "create" to define new named ranges as parameters
/// - Use "delete" to remove obsolete parameters
///
/// Note: Named ranges are Excel's way of creating reusable parameters that can be
/// referenced in formulas and Power Query. They're ideal for configuration values.
/// </summary>
[McpServerToolType]
public static class ExcelParameterTool
{
    /// <summary>
    /// Manage Excel parameters (named ranges) - configuration values and reusable references
    /// </summary>
    [McpServerTool(Name = "excel_parameter")]
    [Description("Manage Excel named ranges as parameters. Supports: list, get, set, update, create, delete. Optional batchId for batch sessions.")]
    public static async Task<string> ExcelParameter(
        [Required]
        [RegularExpression("^(list|get|set|update|create|delete)$")]
        [Description("Action: list, get, set, update, create, delete")]
        string action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [StringLength(255, MinimumLength = 1)]
        [Description("Parameter (named range) name")]
        string? parameterName = null,

        [Description("Parameter value (for set) or cell reference (for create/update, e.g., 'Sheet1!A1')")]
        string? value = null,

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var parameterCommands = new ParameterCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => await ListParametersAsync(parameterCommands, excelPath, batchId),
                "get" => await GetParameterAsync(parameterCommands, excelPath, parameterName, batchId),
                "set" => await SetParameterAsync(parameterCommands, excelPath, parameterName, value, batchId),
                "update" => await UpdateParameterAsync(parameterCommands, excelPath, parameterName, value, batchId),
                "create" => await CreateParameterAsync(parameterCommands, excelPath, parameterName, value, batchId),
                "delete" => await DeleteParameterAsync(parameterCommands, excelPath, parameterName, batchId),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{action}'. Supported: list, get, set, update, create, delete")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action, excelPath);
            throw; // Unreachable but satisfies compiler
        }
    }

    private static async Task<string> ListParametersAsync(ParameterCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListAsync(batch));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Check that the Excel file exists and is accessible",
                "Verify the file path is correct"
            ];
            result.WorkflowHint = "List failed. Ensure the file exists and retry.";
            throw new ModelContextProtocol.McpException($"list failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Use 'get' to retrieve parameter values",
            "Use 'create' to add new parameters",
            "Use 'set' to update existing parameters"
        ];
        result.WorkflowHint = "Parameters listed. Next, get, create, or set values.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetParameterAsync(ParameterCommands commands, string filePath, string? parameterName, string? batchId)
    {
        if (string.IsNullOrEmpty(parameterName))
            throw new ModelContextProtocol.McpException("parameterName is required for get action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetAsync(batch, parameterName));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Check that the parameter name is correct",
                "Use 'list' to see available parameters",
                "Use 'create' to add the parameter if it doesn't exist"
            ];
            result.WorkflowHint = "Get failed. Ensure the parameter exists and retry.";
            throw new ModelContextProtocol.McpException($"get failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Use 'set' to update the parameter value",
            "Use the parameter value in your workflow",
            "Use PowerQuery to reference this parameter"
        ];
        result.WorkflowHint = "Parameter value retrieved. Next, use or update as needed.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetParameterAsync(ParameterCommands commands, string filePath, string? parameterName, string? value, string? batchId)
    {
        if (string.IsNullOrEmpty(parameterName) || value == null)
            throw new ModelContextProtocol.McpException("parameterName and value are required for set action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.SetAsync(batch, parameterName, value));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Check that the parameter exists using 'list'",
                "Use 'create' to add the parameter first",
                "Verify the value format is correct"
            ];
            result.WorkflowHint = "Set failed. Ensure the parameter exists and value is valid.";
            throw new ModelContextProtocol.McpException($"set failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Use 'get' to verify the updated value",
            "Use PowerQuery 'refresh' to update data using new parameter",
            "Verify formulas using this parameter recalculate"
        ];
        result.WorkflowHint = "Parameter updated. Next, verify and refresh dependencies.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateParameterAsync(ParameterCommands commands, string filePath, string? parameterName, string? value, string? batchId)
    {
        if (string.IsNullOrEmpty(parameterName) || string.IsNullOrEmpty(value))
            throw new ModelContextProtocol.McpException("parameterName and value (cell reference) are required for update action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.UpdateAsync(batch, parameterName, value));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Check that the parameter exists using 'list'",
                "Verify the cell reference is valid (e.g., 'Sheet1!A1')",
                "Ensure referenced sheet and cells exist"
            ];
            result.WorkflowHint = "Update failed. Ensure the parameter exists and reference is valid.";
            throw new ModelContextProtocol.McpException($"update failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Use 'get' to verify the new reference",
            "Use 'set' to change the value if needed",
            "Update formulas using this parameter if necessary"
        ];
        result.WorkflowHint = "Parameter reference updated. Next, verify or modify the value.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateParameterAsync(ParameterCommands commands, string filePath, string? parameterName, string? value, string? batchId)
    {
        if (string.IsNullOrEmpty(parameterName) || string.IsNullOrEmpty(value))
            throw new ModelContextProtocol.McpException("parameterName and value (cell reference) are required for create action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.CreateAsync(batch, parameterName, value));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Check that the parameter name doesn't already exist",
                "Verify the cell reference is valid (e.g., 'Sheet1!A1')",
                "Use 'list' to see existing parameters"
            ];
            result.WorkflowHint = "Create failed. Ensure the parameter is unique and reference is valid.";
            throw new ModelContextProtocol.McpException($"create failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Use 'set' to assign an initial value",
            "Use 'get' to verify the parameter",
            "Reference this parameter in PowerQuery or formulas"
        ];
        result.WorkflowHint = "Parameter created. Next, set value and use in workflows.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteParameterAsync(ParameterCommands commands, string filePath, string? parameterName, string? batchId)
    {
        if (string.IsNullOrEmpty(parameterName))
            throw new ModelContextProtocol.McpException("parameterName is required for delete action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.DeleteAsync(batch, parameterName));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Check that the parameter exists",
                "Use 'list' to see available parameters",
                "Verify the parameter name is correct"
            ];
            result.WorkflowHint = "Delete failed. Ensure the parameter exists and name is correct.";
            throw new ModelContextProtocol.McpException($"delete failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Use 'list' to verify the deletion",
            "Update formulas that referenced this parameter",
            "Update PowerQuery code that used this parameter"
        ];
        result.WorkflowHint = "Parameter deleted. Next, update dependent formulas and queries.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
