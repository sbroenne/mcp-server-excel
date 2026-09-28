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
        if (!capabilities.TryGetProperty("provenMethods", out var provenMethods)
            || provenMethods.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidOperationException(
                "The configured helper returned no provenMethods object.");
        }
        var powerQueryListProven = RequiredBoolean(provenMethods, "powerQueryList");
        var powerQueryMutationProven = RequiredBoolean(provenMethods, "powerQueryMutation");

        var result = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in supportedActions.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String
                || item.GetString() is not { } action)
            {
                throw new InvalidOperationException(
                    "The configured helper returned an invalid supported action.");
            }
            var isProven = action switch
            {
                "powerquery.list" or "powerquery.view" => powerQueryListProven,
                "powerquery.create"
                    or "powerquery.update"
                    or "powerquery.rename"
                    or "powerquery.delete" => powerQueryMutationProven,
                _ => false
            };
            if (isProven)
            {
                result.Add(action);
            }
        }
        return result;
    }

    private static bool RequiredBoolean(JsonElement value, string propertyName)
    {
        if (!value.TryGetProperty(propertyName, out var property)
            || property.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
        {
            throw new InvalidOperationException(
                $"The configured helper returned no valid provenMethods.{propertyName} value.");
        }
        return property.GetBoolean();
    }
}
