using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacVbaHelperCapabilities
{
    private const string CandidateActionsEnvironmentVariable =
        "EXCELMCP_MAC_VBA_CANDIDATE_ACTIONS";
    private readonly HashSet<string> _supportedActions;
    private readonly HashSet<string> _candidateActions;

    private MacVbaHelperCapabilities(
        HashSet<string> supportedActions,
        HashSet<string> candidateActions,
        bool projectReadable,
        bool listViewProven,
        bool mutationProven,
        bool runProven)
    {
        _supportedActions = supportedActions;
        _candidateActions = candidateActions;
        ProjectReadable = projectReadable;
        ListViewProven = listViewProven;
        MutationProven = mutationProven;
        RunProven = runProven;
    }

    public bool ProjectReadable { get; }

    public bool ListViewProven { get; }

    public bool MutationProven { get; }

    public bool RunProven { get; }

    public static IReadOnlySet<string> GetExplicitOptIn()
    {
        var configured = Environment.GetEnvironmentVariable(
            CandidateActionsEnvironmentVariable);
        if (string.IsNullOrWhiteSpace(configured))
        {
            return new HashSet<string>(StringComparer.Ordinal);
        }
        return configured
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(action => action is
                "vba.list"
                or "vba.view"
                or "vba.import"
                or "vba.update"
                or "vba.delete"
                or "vba.run")
            .ToHashSet(StringComparer.Ordinal);
    }

    public static MacVbaHelperCapabilities Parse(
        JsonElement root,
        IReadOnlySet<string> candidateActions)
    {
        var supportedActions = new HashSet<string>(StringComparer.Ordinal);
        if (root.TryGetProperty("supportedActions", out var actions)
            && actions.ValueKind == JsonValueKind.Array)
        {
            foreach (var action in actions.EnumerateArray())
            {
                if (action.ValueKind == JsonValueKind.String
                    && action.GetString() is { } value)
                {
                    supportedActions.Add(value);
                }
            }
        }

        var projectReadable = root.TryGetProperty("trustReadiness", out var trust)
            && trust.ValueKind == JsonValueKind.Object
            && trust.TryGetProperty("vbaProjectReadable", out var readable)
            && readable.ValueKind == JsonValueKind.True;
        var proof = root.TryGetProperty("provenMethods", out var methods)
            && methods.ValueKind == JsonValueKind.Object
                ? methods
                : default;

        return new(
            supportedActions,
            new HashSet<string>(candidateActions, StringComparer.Ordinal),
            projectReadable,
            Proven(proof, "vbaListView"),
            Proven(proof, "vbaMutation"),
            Proven(proof, "vbaRun"));
    }

    public bool IsAvailable(string action)
    {
        if (!_supportedActions.Contains(action))
        {
            return false;
        }
        if (_candidateActions.Contains(action))
        {
            return true;
        }
        return action switch
        {
            "vba.list" or "vba.view" => ListViewProven,
            "vba.import" or "vba.update" or "vba.delete" => MutationProven,
            "vba.run" => RunProven,
            _ => false
        };
    }

    private static bool Proven(JsonElement proof, string propertyName) =>
        proof.ValueKind == JsonValueKind.Object
        && proof.TryGetProperty(propertyName, out var value)
        && value.ValueKind == JsonValueKind.True;
}
