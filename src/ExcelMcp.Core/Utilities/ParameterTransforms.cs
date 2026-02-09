using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Commands.Range;

namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>
/// Shared parameter transformation utilities used by MCP, CLI, and generated code.
/// These provide consistent handling of common patterns across all entry points.
/// </summary>
public static class ParameterTransforms
{
    private static readonly JsonSerializerOptions s_jsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    // === JSON List Parsing ===

    /// <summary>
    /// Parses a JSON array string into a list of strings.
    /// Returns null if input is null/empty.
    /// </summary>
    /// <param name="json">JSON array string, e.g. '["value1","value2"]'</param>
    /// <param name="parameterName">Parameter name for error messages</param>
    /// <returns>Parsed list or null</returns>
    /// <exception cref="ArgumentException">Thrown when JSON is invalid</exception>
    public static List<string>? ParseJsonList(string? json, string parameterName = "value")
    {
        if (string.IsNullOrWhiteSpace(json))
            return null;

        try
        {
            return JsonSerializer.Deserialize<List<string>>(json!, s_jsonOptions);
        }
        catch (JsonException ex)
        {
            throw new ArgumentException(
                $"Invalid {parameterName} JSON: {ex.Message}. Expected: '[\"value1\",\"value2\"]'",
                parameterName);
        }
    }

    /// <summary>
    /// Parses a JSON array string into a list of strings, with single-item fallback.
    /// If the string is not valid JSON, treats it as a single item.
    /// Returns null if input is null/empty.
    /// </summary>
    /// <param name="json">JSON array string or single value</param>
    /// <returns>Parsed list, single-item list, or null</returns>
    public static List<string>? ParseJsonListOrSingle(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
            return null;

        try
        {
            return JsonSerializer.Deserialize<List<string>>(json!, s_jsonOptions);
        }
        catch (JsonException)
        {
            // If parsing fails, treat as single item
            return [json!];
        }
    }

    /// <summary>
    /// Deserializes a JSON string into a typed object.
    /// Returns default if input is null/empty.
    /// </summary>
    /// <typeparam name="T">Target type</typeparam>
    /// <param name="json">JSON string</param>
    /// <param name="parameterName">Parameter name for error messages</param>
    /// <returns>Deserialized object or default</returns>
    /// <exception cref="ArgumentException">Thrown when JSON is invalid</exception>
    public static T? DeserializeJson<T>(string? json, string parameterName = "value") where T : class
    {
        if (string.IsNullOrWhiteSpace(json))
            return null;

        try
        {
            return JsonSerializer.Deserialize<T>(json!, s_jsonOptions);
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid {parameterName} JSON: {ex.Message}", parameterName);
        }
    }

    // === CSV Parsing ===

    /// <summary>
    /// Splits a comma-separated string into a trimmed string array.
    /// Returns null if input is null/empty.
    /// </summary>
    /// <param name="csv">Comma-separated values</param>
    /// <returns>Array of trimmed values, or null</returns>
    public static string[]? SplitCsvParameters(string? csv)
    {
        if (string.IsNullOrWhiteSpace(csv))
            return null;

        return csv.Split(',', StringSplitOptions.RemoveEmptyEntries)
                  .Select(p => p.Trim())
                  .ToArray();
    }

    /// <summary>
    /// Parses multi-line CSV text into a 2D list of values for table operations.
    /// Each line becomes a row, comma-separated values become cells.
    /// Quoted values have surrounding quotes stripped.
    /// Returns null if input is null/empty.
    /// </summary>
    /// <param name="csvData">Multi-line CSV text</param>
    /// <returns>2D list of values, or null</returns>
    public static List<List<object?>>? ParseCsvToRows(string? csvData)
    {
        if (string.IsNullOrWhiteSpace(csvData))
            return null;

        var lines = csvData!.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

        return lines.Select(line =>
        {
            var values = line.Split(',');
            return values.Select(value =>
            {
                var trimmed = value.Trim().Trim('"');
                return string.IsNullOrEmpty(trimmed) ? null : (object?)trimmed;
            }).ToList();
        }).ToList();
    }

    // === Options Object Construction ===

    /// <summary>
    /// Resolves values from either an inline 2D array or a file path.
    /// Supports JSON files (2D array format) and CSV files (rows/columns).
    /// File format is auto-detected from extension (.json → JSON, anything else → CSV).
    /// </summary>
    /// <param name="values">Inline 2D array of values (may be null if file is provided)</param>
    /// <param name="valuesFile">Path to JSON or CSV file containing values</param>
    /// <param name="parameterName">Parameter name for error messages</param>
    /// <returns>Resolved 2D array of values</returns>
    /// <exception cref="ArgumentException">Neither values nor valuesFile provided</exception>
    /// <exception cref="FileNotFoundException">File not found</exception>
    public static List<List<object?>> ResolveValuesOrFile(List<List<object?>>? values, string? valuesFile, string parameterName = "values")
    {
        if (values != null && values.Count > 0)
            return values;

        if (string.IsNullOrWhiteSpace(valuesFile))
            throw new ArgumentException($"Either {parameterName} or {parameterName}File must be provided for set-values action", parameterName);

        if (!File.Exists(valuesFile))
            throw new FileNotFoundException($"Values file not found: {valuesFile}", valuesFile);

        var content = File.ReadAllText(valuesFile);

        if (valuesFile.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var parsed = JsonSerializer.Deserialize<List<List<object?>>>(content, s_jsonOptions);
                return parsed ?? throw new ArgumentException($"JSON file '{valuesFile}' deserialized to null");
            }
            catch (JsonException ex)
            {
                throw new ArgumentException(
                    $"Invalid JSON in values file '{valuesFile}': {ex.Message}. Expected 2D array: [[1,2],[3,4]]",
                    parameterName);
            }
        }

        // Default: treat as CSV
        var csvRows = ParseCsvToRows(content);
        if (csvRows == null || csvRows.Count == 0)
            throw new ArgumentException($"Values file '{valuesFile}' is empty or contains no parseable data");

        return csvRows;
    }

    /// <summary>
    /// Resolves string formulas from either an inline 2D array or a JSON file path.
    /// </summary>
    /// <param name="formulas">Inline 2D array of formulas (may be null if file is provided)</param>
    /// <param name="formulasFile">Path to JSON file containing formulas</param>
    /// <param name="parameterName">Parameter name for error messages</param>
    /// <returns>Resolved 2D array of formulas</returns>
    public static List<List<string>> ResolveFormulasOrFile(List<List<string>>? formulas, string? formulasFile, string parameterName = "formulas")
    {
        if (formulas != null && formulas.Count > 0)
            return formulas;

        if (string.IsNullOrWhiteSpace(formulasFile))
            throw new ArgumentException($"Either {parameterName} or {parameterName}File must be provided for set-formulas action", parameterName);

        if (!File.Exists(formulasFile))
            throw new FileNotFoundException($"Formulas file not found: {formulasFile}", formulasFile);

        var content = File.ReadAllText(formulasFile);
        try
        {
            var parsed = JsonSerializer.Deserialize<List<List<string>>>(content, s_jsonOptions);
            return parsed ?? throw new ArgumentException($"JSON file '{formulasFile}' deserialized to null");
        }
        catch (JsonException ex)
        {
            throw new ArgumentException(
                $"Invalid JSON in formulas file '{formulasFile}': {ex.Message}. Expected 2D array: [[\"=A1+B1\"],[\"=SUM(A:A)\"]]",
                parameterName);
        }
    }

    // === Options Object Construction ===

    /// <summary>
    /// Builds a FindOptions object from individual boolean parameters with defaults.
    /// </summary>
    public static FindOptions BuildFindOptions(
        bool? matchCase = null,
        bool? matchEntireCell = null,
        bool? searchFormulas = null,
        bool? searchValues = null)
    {
        return new FindOptions
        {
            MatchCase = matchCase ?? false,
            MatchEntireCell = matchEntireCell ?? false,
            SearchFormulas = searchFormulas ?? true,
            SearchValues = searchValues ?? true
        };
    }

    /// <summary>
    /// Builds a ReplaceOptions object from individual boolean parameters with defaults.
    /// </summary>
    public static ReplaceOptions BuildReplaceOptions(
        bool? matchCase = null,
        bool? matchEntireCell = null,
        bool? replaceAll = null)
    {
        return new ReplaceOptions
        {
            MatchCase = matchCase ?? false,
            MatchEntireCell = matchEntireCell ?? false,
            ReplaceAll = replaceAll ?? true
        };
    }

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
