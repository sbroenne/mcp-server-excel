using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.Core.Models.Validation;

/// <summary>
/// Defines an action with its CLI command, MCP details, and parameter validation
/// Provides single source of truth for action metadata across all layers
/// </summary>
public class ActionDefinition
{
    /// <summary>
    /// Domain/category of action (e.g., "PowerQuery", "Parameter", "Table")
    /// </summary>
    public string Domain { get; init; } = "";

    /// <summary>
    /// Generic action name (e.g., "list", "view", "create")
    /// </summary>
    public string Action { get; init; } = "";

    /// <summary>
    /// CLI command name (e.g., "pq-list", "param-create", "table-info")
    /// </summary>
    public string CliCommand { get; init; } = "";

    /// <summary>
    /// MCP action name (e.g., "list", "view", "create")
    /// </summary>
    public string McpAction { get; init; } = "";

    /// <summary>
    /// MCP tool name (e.g., "excel_powerquery", "excel_parameter")
    /// </summary>
    public string McpTool { get; init; } = "";

    /// <summary>
    /// Parameter definitions for this action
    /// </summary>
    public ParameterDefinition[] Parameters { get; init; } = Array.Empty<ParameterDefinition>();

    /// <summary>
    /// Description of what this action does
    /// </summary>
    public string? Description { get; init; }

    /// <summary>
    /// Validates parameters for this action
    /// </summary>
    public ValidationResult ValidateParameters(Dictionary<string, object?> parameters)
    {
        foreach (var paramDef in Parameters)
        {
            parameters.TryGetValue(paramDef.Name, out var value);
            var result = paramDef.Validate(value);
            if (!result.IsValid)
            {
                return result;
            }
        }

        return ValidationResult.Success();
    }

    /// <summary>
    /// Validates CLI arguments for this action
    /// </summary>
    public ValidationResult ValidateCliArgs(string[] args)
    {
        // First arg is command name, so required params start at index 1
        int requiredCount = Parameters.Count(p => p.Required);
        int providedCount = args.Length - 1; // Subtract command name

        if (providedCount < requiredCount)
        {
            return ValidationResult.Failure("args",
                $"Usage: excelcli {CliCommand} {GetCliUsage()}");
        }

        // Validate each parameter
        int argIndex = 1;
        foreach (var paramDef in Parameters.Where(p => p.Required))
        {
            if (argIndex < args.Length)
            {
                var result = paramDef.Validate(args[argIndex]);
                if (!result.IsValid)
                {
                    return result;
                }
                argIndex++;
            }
        }

        return ValidationResult.Success();
    }

    /// <summary>
    /// Gets CLI usage string
    /// </summary>
    public string GetCliUsage()
    {
        var parts = new List<string>();

        foreach (var param in Parameters)
        {
            if (param.Required)
            {
                parts.Add($"<{param.Name}>");
            }
            else
            {
                parts.Add($"[{param.Name}]");
            }
        }

        return string.Join(" ", parts);
    }

    /// <summary>
    /// Gets regex pattern for MCP action validation
    /// </summary>
    public string GetMcpActionPattern()
    {
        return $"^{Regex.Escape(McpAction)}$";
    }
}
