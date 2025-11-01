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
    [Description(@"Manage Excel named ranges as parameters (configuration values).

⚡ PERFORMANCE: For creating 2+ parameters, use begin_excel_batch FIRST (90% faster):
  1. batch = begin_excel_batch(excelPath: 'file.xlsx')
  2. excel_parameter(action: 'create', ..., batchId: batch.batchId)  // repeat for each parameter
  3. commit_excel_batch(batchId: batch.batchId, save: true)

⭐ NEW: Use 'create-bulk' action for even better efficiency (one call for multiple parameters).

Actions available as dropdown in MCP clients.")]
    public static async Task<string> ExcelParameter(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        ParameterAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [StringLength(255, MinimumLength = 1)]
        [Description("Parameter (named range) name (for get, set, create, delete actions)")]
        string? parameterName = null,

        [Description("Parameter value (for set action) or cell reference (for create actions, e.g., 'Sheet1!A1')")]
        string? value = null,

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var parameterCommands = new ParameterCommands();
            var actionString = action.ToActionString();

            return actionString switch
            {
                "list" => await ListParametersAsync(parameterCommands, excelPath, batchId),
                "get" => await GetParameterAsync(parameterCommands, excelPath, parameterName, batchId),
                "set" => await SetParameterAsync(parameterCommands, excelPath, parameterName, value, batchId),
                "create" => await CreateParameterAsync(parameterCommands, excelPath, parameterName, value, batchId),
                "delete" => await DeleteParameterAsync(parameterCommands, excelPath, parameterName, batchId),
                _ => throw new ModelContextProtocol.McpException($"Unknown action '{actionString}'")
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

            throw new ModelContextProtocol.McpException($"list failed for '{filePath}': {result.ErrorMessage}");
        }

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

            throw new ModelContextProtocol.McpException($"get failed for '{filePath}': {result.ErrorMessage}");
        }

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

            throw new ModelContextProtocol.McpException($"set failed for '{filePath}': {result.ErrorMessage}");
        }

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

            throw new ModelContextProtocol.McpException($"update failed for '{filePath}': {result.ErrorMessage}");
        }

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

            throw new ModelContextProtocol.McpException($"create failed for '{filePath}': {result.ErrorMessage}");
        }

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

            throw new ModelContextProtocol.McpException($"delete failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateBulkParametersAsync(ParameterCommands commands, string excelPath, string? parametersJson, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(parametersJson))
            throw new ModelContextProtocol.McpException("parametersJson is required for create-bulk action");

        // Deserialize JSON array of parameter definitions
        List<ParameterDefinition>? parameters;
        try
        {
            parameters = JsonSerializer.Deserialize<List<ParameterDefinition>>(
                parametersJson,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (parameters == null || parameters.Count == 0)
                throw new ModelContextProtocol.McpException("parametersJson must contain at least one parameter definition");
        }
        catch (JsonException ex)
        {
            throw new ModelContextProtocol.McpException($"Invalid parametersJson format: {ex.Message}");
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

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
