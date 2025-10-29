using Sbroenne.ExcelMcp.Core.Models.Validation;

namespace Sbroenne.ExcelMcp.CLI.Validation;

/// <summary>
/// Maps domain ActionDefinitions to CLI-specific metadata
/// This adapter lives in CLI layer, keeping validation layer client-agnostic
/// </summary>
public static class CliActionMapper
{
    /// <summary>
    /// CLI command prefixes by domain
    /// </summary>
    private static readonly Dictionary<string, string> CliCommandPrefixes = new()
    {
        ["PowerQuery"] = "pq",
        ["Parameter"] = "param",
        ["Table"] = "table",
        ["DataModel"] = "dm",
        ["VBA"] = "vba",
        ["Connection"] = "conn",
        ["Worksheet"] = "sheet",
        ["Range"] = "range",
        ["File"] = "file"
    };

    /// <summary>
    /// Gets CLI command name for an action definition
    /// </summary>
    public static string GetCliCommandName(ActionDefinition action)
    {
        string prefix = CliCommandPrefixes.TryGetValue(action.Domain, out var p) 
            ? p 
            : action.Domain.ToLowerInvariant();
        return $"{prefix}-{action.Name}";
    }

    /// <summary>
    /// Gets CLI usage string for an action
    /// </summary>
    public static string GetCliUsage(ActionDefinition action)
    {
        var parts = new List<string>();

        foreach (var param in action.Parameters)
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
    /// Validates CLI arguments for an action
    /// </summary>
    public static ValidationResult ValidateCliArgs(ActionDefinition action, string[] args)
    {
        // First arg is command name, so required params start at index 1
        int requiredCount = action.Parameters.Count(p => p.Required);
        int providedCount = args.Length - 1; // Subtract command name

        if (providedCount < requiredCount)
        {
            return ValidationResult.Failure("args",
                $"Usage: excelcli {GetCliCommandName(action)} {GetCliUsage(action)}");
        }

        // Validate each parameter
        int argIndex = 1;
        foreach (var paramDef in action.Parameters.Where(p => p.Required))
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
}
