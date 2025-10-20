using Sbroenne.ExcelMcp.Core.Commands;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel parameter (named range) management tool for MCP server.
/// Handles named ranges as configuration parameters for Excel automation.
/// 
/// LLM Usage Patterns:
/// - Use "list" to see all named ranges (parameters) in a workbook
/// - Use "get" to retrieve parameter values for configuration
/// - Use "set" to update parameter values for dynamic behavior
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
    [Description("Manage Excel named ranges as parameters. Supports: list, get, set, create, delete.")]
    public static string ExcelParameter(
        [Required]
        [RegularExpression("^(list|get|set|create|delete)$")]
        [Description("Action: list, get, set, create, delete")] 
        string action,
        
        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")] 
        string excelPath,
        
        [StringLength(255, MinimumLength = 1)]
        [Description("Parameter (named range) name")] 
        string? parameterName = null,
        
        [Description("Parameter value (for set) or cell reference (for create, e.g., 'Sheet1!A1')")] 
        string? value = null)
    {
        try
        {
            var parameterCommands = new ParameterCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ListParameters(parameterCommands, excelPath),
                "get" => GetParameter(parameterCommands, excelPath, parameterName),
                "set" => SetParameter(parameterCommands, excelPath, parameterName, value),
                "create" => CreateParameter(parameterCommands, excelPath, parameterName, value),
                "delete" => DeleteParameter(parameterCommands, excelPath, parameterName),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{action}'. Supported: list, get, set, create, delete")
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

    private static string ListParameters(ParameterCommands commands, string filePath)
    {
        var result = commands.List(filePath);
        
        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"list failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string GetParameter(ParameterCommands commands, string filePath, string? parameterName)
    {
        if (string.IsNullOrEmpty(parameterName))
            throw new ModelContextProtocol.McpException("parameterName is required for get action");

        var result = commands.Get(filePath, parameterName);
        
        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetParameter(ParameterCommands commands, string filePath, string? parameterName, string? value)
    {
        if (string.IsNullOrEmpty(parameterName) || value == null)
            throw new ModelContextProtocol.McpException("parameterName and value are required for set action");

        var result = commands.Set(filePath, parameterName, value);
        
        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string CreateParameter(ParameterCommands commands, string filePath, string? parameterName, string? value)
    {
        if (string.IsNullOrEmpty(parameterName) || string.IsNullOrEmpty(value))
            throw new ModelContextProtocol.McpException("parameterName and value (cell reference) are required for create action");

        var result = commands.Create(filePath, parameterName, value);
        
        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"create failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteParameter(ParameterCommands commands, string filePath, string? parameterName)
    {
        if (string.IsNullOrEmpty(parameterName))
            throw new ModelContextProtocol.McpException("parameterName is required for delete action");

        var result = commands.Delete(filePath, parameterName);
        
        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"delete failed for '{filePath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}