using Sbroenne.ExcelMcp.Core.Commands;
using ModelContextProtocol.Server;
using System.ComponentModel;
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
public static class ExcelParameterTool
{
    /// <summary>
    /// Manage Excel parameters (named ranges) - configuration values and reusable references
    /// </summary>
    [McpServerTool(Name = "excel_parameter")]
    [Description("Manage Excel named ranges as parameters. Supports: list, get, set, create, delete.")]
    public static string ExcelParameter(
        [Description("Action: list, get, set, create, delete")] string action,
        [Description("Excel file path (.xlsx or .xlsm)")] string filePath,
        [Description("Parameter (named range) name")] string? parameterName = null,
        [Description("Parameter value (for set) or cell reference (for create, e.g., 'Sheet1!A1')")] string? value = null)
    {
        try
        {
            var parameterCommands = new ParameterCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ListParameters(parameterCommands, filePath),
                "get" => GetParameter(parameterCommands, filePath, parameterName),
                "set" => SetParameter(parameterCommands, filePath, parameterName, value),
                "create" => CreateParameter(parameterCommands, filePath, parameterName, value),
                "delete" => DeleteParameter(parameterCommands, filePath, parameterName),
                _ => ExcelToolsBase.CreateUnknownActionError(action, "list", "get", "set", "create", "delete")
            };
        }
        catch (Exception ex)
        {
            return ExcelToolsBase.CreateExceptionError(ex, action, filePath);
        }
    }

    private static string ListParameters(ParameterCommands commands, string filePath)
    {
        var result = commands.List(filePath);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string GetParameter(ParameterCommands commands, string filePath, string? parameterName)
    {
        if (string.IsNullOrEmpty(parameterName))
            return JsonSerializer.Serialize(new { error = "parameterName is required for get action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Get(filePath, parameterName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetParameter(ParameterCommands commands, string filePath, string? parameterName, string? value)
    {
        if (string.IsNullOrEmpty(parameterName) || value == null)
            return JsonSerializer.Serialize(new { error = "parameterName and value are required for set action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Set(filePath, parameterName, value);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string CreateParameter(ParameterCommands commands, string filePath, string? parameterName, string? value)
    {
        if (string.IsNullOrEmpty(parameterName) || string.IsNullOrEmpty(value))
            return JsonSerializer.Serialize(new { error = "parameterName and value (cell reference) are required for create action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Create(filePath, parameterName, value);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteParameter(ParameterCommands commands, string filePath, string? parameterName)
    {
        if (string.IsNullOrEmpty(parameterName))
            return JsonSerializer.Serialize(new { error = "parameterName is required for delete action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Delete(filePath, parameterName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}