using System.Text.Json;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Classifies a service result payload the same way the MCP Server does: a JSON object
/// whose root reports <c>success: false</c> or <c>isError: true</c> is a failed operation,
/// even when the service delivered it without an exception.
/// </summary>
internal static class ServiceResultOutcome
{
    internal static bool IsNegative(string? resultJson) =>
        TryReadNegative(resultJson, out _, out _);

    internal static bool TryReadNegative(
        string? resultJson,
        out string? errorMessage,
        out string? errorCategory)
    {
        errorMessage = null;
        errorCategory = null;
        if (string.IsNullOrWhiteSpace(resultJson))
        {
            return false;
        }

        try
        {
            using var document = JsonDocument.Parse(resultJson);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            var negative =
                (root.TryGetProperty("success", out var success) && success.ValueKind == JsonValueKind.False) ||
                (root.TryGetProperty("isError", out var isError) && isError.ValueKind == JsonValueKind.True);
            if (!negative)
            {
                return false;
            }

            errorMessage = ReadString(root, "errorMessage");
            errorCategory = ReadString(root, "errorCategory");
            return true;
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static string? ReadString(JsonElement root, string name) =>
        root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;
}
