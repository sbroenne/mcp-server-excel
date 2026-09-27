using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacPowerQueryHelperCapabilities
{
    public static IReadOnlySet<string> Parse(JsonElement capabilities)
    {
        if (!capabilities.TryGetProperty("supportedActions", out var supportedActions)
            || supportedActions.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException(
                "The configured helper returned no supportedActions array.");
        }

        var result = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in supportedActions.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String
                || item.GetString() is not { } action)
            {
                throw new InvalidOperationException(
                    "The configured helper returned an invalid supported action.");
            }
            if (action.StartsWith("powerquery.", StringComparison.Ordinal))
            {
                result.Add(action);
            }
        }
        return result;
    }
}
