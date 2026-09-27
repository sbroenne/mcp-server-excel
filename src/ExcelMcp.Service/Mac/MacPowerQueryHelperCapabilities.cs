using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacPowerQueryHelperCapabilities
{
    internal const string CandidateActionsEnvironmentVariable =
        "EXCELMCP_MAC_POWERQUERY_CANDIDATE_ACTIONS";

    private static readonly Dictionary<string, string> ProofProperties =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["powerquery.list"] = "powerQueryList",
            ["powerquery.view"] = "powerQueryList",
            ["powerquery.create"] = "powerQueryCreate",
            ["powerquery.update"] = "powerQueryUpdate",
            ["powerquery.rename"] = "powerQueryRename",
            ["powerquery.delete"] = "powerQueryDelete",
            ["powerquery.refresh"] = "powerQueryRefresh",
            ["powerquery.refresh-all"] = "powerQueryRefreshAll",
            ["powerquery.load-to"] = "powerQueryLoadTo",
            ["powerquery.unload"] = "powerQueryUnload",
            ["powerquery.evaluate"] = "powerQueryEvaluate"
        };

    public static IReadOnlySet<string> GetExplicitOptIn() =>
        ParseExplicitOptIn(
            Environment.GetEnvironmentVariable(CandidateActionsEnvironmentVariable));

    internal static IReadOnlySet<string> ParseExplicitOptIn(string? configuredActions)
    {
        if (string.IsNullOrWhiteSpace(configuredActions))
        {
            return new HashSet<string>(StringComparer.Ordinal);
        }

        var result = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in configuredActions.Split(
                     ',',
                     StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (!ProofProperties.ContainsKey(item))
            {
                throw new InvalidOperationException(
                    $"{CandidateActionsEnvironmentVariable} contains unsupported action '{item}'.");
            }
            result.Add(item);
        }
        return result;
    }

    public static IReadOnlySet<string> Parse(
        JsonElement capabilities,
        IReadOnlySet<string>? explicitlyEnabledActions = null)
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
        explicitlyEnabledActions ??= new HashSet<string>(StringComparer.Ordinal);

        var result = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in supportedActions.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String
                || item.GetString() is not { } action)
            {
                throw new InvalidOperationException(
                    "The configured helper returned an invalid supported action.");
            }
            if (ProofProperties.TryGetValue(action, out var proofProperty)
                && (RequiredBoolean(provenMethods, proofProperty)
                    || explicitlyEnabledActions.Contains(action)))
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
