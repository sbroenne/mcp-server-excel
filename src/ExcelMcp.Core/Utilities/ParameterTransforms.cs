namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>
/// Shared parameter transformation utilities used by MCP, CLI, and generated code.
/// These provide consistent handling of common patterns across all entry points.
/// </summary>
public static class ParameterTransforms
{
    /// <summary>
    /// Resolves a value that can come from either a direct string or a file path.
    /// If filePath is provided and exists, reads file content. Otherwise returns directValue.
    /// </summary>
    /// <param name="directValue">The direct string value (e.g., M code inline)</param>
    /// <param name="filePath">Optional path to a file containing the value</param>
    /// <returns>The resolved value (file content or direct value)</returns>
    public static string? ResolveFileOrValue(string? directValue, string? filePath)
    {
        if (!string.IsNullOrWhiteSpace(filePath))
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found: {filePath}", filePath);
            }
            return File.ReadAllText(filePath);
        }
        return directValue;
    }

    /// <summary>
    /// Parses a string load destination to the PowerQueryLoadMode enum.
    /// </summary>
    /// <param name="loadDestination">String value: "worksheet", "data-model", "both", "connection-only"</param>
    /// <returns>The corresponding PowerQueryLoadMode enum value</returns>
    public static Models.PowerQueryLoadMode ParseLoadMode(string? loadDestination)
    {
        return loadDestination?.ToLowerInvariant() switch
        {
            "worksheet" or "table" => Models.PowerQueryLoadMode.LoadToTable,
            "data-model" or "datamodel" => Models.PowerQueryLoadMode.LoadToDataModel,
            "both" => Models.PowerQueryLoadMode.LoadToBoth,
            "connection-only" or "connectiononly" => Models.PowerQueryLoadMode.ConnectionOnly,
            _ => Models.PowerQueryLoadMode.LoadToTable
        };
    }

    /// <summary>
    /// Validates that a required parameter is not null or empty.
    /// </summary>
    /// <param name="value">The parameter value to validate</param>
    /// <param name="parameterName">Name of the parameter for error messages</param>
    /// <param name="actionName">Name of the action for error messages</param>
    /// <exception cref="ArgumentException">Thrown when value is null or empty</exception>
    public static void RequireNotEmpty(string? value, string parameterName, string actionName)
    {
        if (string.IsNullOrEmpty(value))
        {
            throw new ArgumentException($"{parameterName} is required for {actionName} action", parameterName);
        }
    }

    /// <summary>
    /// Validates that a required parameter is not null or empty, returning the value if valid.
    /// </summary>
    /// <param name="value">The parameter value to validate</param>
    /// <param name="parameterName">Name of the parameter for error messages</param>
    /// <param name="actionName">Name of the action for error messages</param>
    /// <returns>The validated non-null value</returns>
    /// <exception cref="ArgumentException">Thrown when value is null or empty</exception>
    public static string RequireNotEmptyReturn(string? value, string parameterName, string actionName)
    {
        RequireNotEmpty(value, parameterName, actionName);
        return value!;
    }
}
