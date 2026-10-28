using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Marks nested operation results whose failures make the enclosing operation fail.</summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class ResultOutcomeAttribute : Attribute;

/// <summary>Outcome information generated from an operation's source return type.</summary>
public sealed record ResultContract(
    bool ReadOnly,
    string[] OutcomePaths);

/// <summary>One interpreted operation outcome, independent of its delivery protocol.</summary>
public sealed class ResultJson
{
    /// <summary>Whether the operation succeeded, rather than merely being delivered.</summary>
    public bool Success { get; private set; }

    /// <summary>Business data, with root outcome metadata removed.</summary>
    public JsonNode? Data { get; private set; }
    private bool HasData { get; set; }

    /// <summary>Warnings reported by the operation.</summary>
    public JsonNode? Warnings { get; private set; }

    /// <summary>Structured failure information.</summary>
    public JsonObject? Error { get; private set; }

    /// <summary>Exact retained state reported by the execution owner.</summary>
    public JsonNode? PartialState { get; private set; }

    /// <summary>Serialization preserves false, zero, and null business values.</summary>
    public static JsonSerializerOptions Options { get; } = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    /// <summary>Reads an operation result without losing an inner negative outcome.</summary>
    public static ResultJson Read(
        string? json,
        string command,
        bool delivered = true,
        string? message = null,
        string? category = null,
        string? sessionId = null,
        string? exceptionType = null,
        string? hresult = null,
        string? innerError = null,
        string? partialState = null,
        ResultContract? contract = null)
    {
        var data = json is null ? null : JsonNode.Parse(json);
        var root = data as JsonObject;
        var failed = !delivered || IsNegative(root) ||
            command == "session.test" && root?["canOpen"] is JsonValue canOpen &&
            canOpen.TryGetValue<bool>(out var openable) && !openable;
        JsonObject? nestedFailure = null;
        if (contract is not null)
        {
            foreach (var path in contract.OutcomePaths)
            {
                nestedFailure ??= EnumerateAt(data, path.Split('.')).OfType<JsonObject>().FirstOrDefault(IsNegative);
                failed |= nestedFailure is not null;
            }
        }

        var result = new ResultJson
        {
            Success = !failed,
            Data = data,
            HasData = json is not null,
            Warnings = ReadWarnings(root),
            PartialState = partialState is not null
                ? JsonNode.Parse(partialState)
                : root?["partialState"]?.DeepClone()
        };
        if (failed)
        {
            result.Error = new JsonObject
            {
                ["category"] = category ?? ReadString(root, "errorCategory") ??
                    ReadString(nestedFailure, "errorCategory") ?? "OperationFailed",
                ["message"] = message ?? ReadString(root, "errorMessage") ??
                    ReadString(nestedFailure, "errorMessage") ?? ReadString(root, "message") ??
                    "The operation reported a failure.",
                ["command"] = command
            };
            Add(result.Error, "sessionId", sessionId ?? ReadString(root, "sessionId") ?? ReadString(root, "workbook_session_id"));
            Add(result.Error, "exceptionType", exceptionType ?? ReadString(root, "exceptionType") ?? ReadString(nestedFailure, "exceptionType"));
            Add(result.Error, "hresult", hresult ?? ReadString(root, "hresult") ?? ReadString(nestedFailure, "hresult"));
            Add(result.Error, "innerError", innerError ?? ReadString(root, "innerError") ?? ReadString(nestedFailure, "innerError"));
        }

        if (root is not null)
        {
            foreach (var key in new[] { "success", "isError", "errorMessage", "errorCategory",
                "exceptionType", "hresult", "innerError", "warnings", "warning", "partialState" })
                root.Remove(key);
            if (ReadString(root, "command") == command)
                root.Remove("command");
            if (failed)
            {
                root.Remove("sessionId");
                root.Remove("workbook_session_id");
            }
            if (delivered && !failed && !command.StartsWith("session.", StringComparison.Ordinal) &&
                !command.Equals("chart.export-image", StringComparison.OrdinalIgnoreCase))
                root.Remove("filePath");
            if (root["action"] is JsonValue action &&
                action.TryGetValue<string>(out var actionName) &&
                command.EndsWith("." + actionName, StringComparison.Ordinal))
                root.Remove("action");
            if (root.Count == 0)
            {
                result.Data = null;
                result.HasData = false;
            }
            if (!failed && command.Equals("powerquery.list", StringComparison.OrdinalIgnoreCase) &&
                root["queries"] is JsonArray queries)
                foreach (var query in queries.OfType<JsonObject>())
                    query.Remove("formulaPreview");
        }
        return result;
    }

    /// <summary>Adds a later failure without changing the completed operation's data shape.</summary>
    public ResultJson WithFailure(string command, Exception exception, string category, string message,
        string? partialState = null)
    {
        var failure = Read(null, command, delivered: false, message: message, category: category,
            exceptionType: exception.GetType().Name, hresult: $"0x{exception.HResult:X8}",
            innerError: exception.InnerException?.Message, partialState: partialState);
        failure.Data = Data?.DeepClone();
        failure.HasData = HasData;
        failure.Warnings = Warnings?.DeepClone();
        failure.PartialState ??= PartialState?.DeepClone();
        return failure;
    }

    /// <summary>Returns the business payload in MCP's existing object form.</summary>
    public string ToToolJson()
    {
        var root = Data is JsonObject data
            ? (JsonObject)data.DeepClone()
            : HasData ? new JsonObject { ["result"] = Data?.DeepClone() } : new JsonObject();
        root["success"] = Success;
        if (Warnings is not null)
            root["warnings"] = Warnings.DeepClone();
        if (PartialState is not null)
            root["partialState"] = PartialState.DeepClone();
        if (Error is not null)
        {
            root["errorMessage"] = Error["message"]?.DeepClone();
            root["errorCategory"] = Error["category"]?.DeepClone();
            foreach (var property in Error.Where(property => property.Key is not ("category" or "message")))
                root[property.Key == "sessionId" ? "workbook_session_id" : property.Key] = property.Value?.DeepClone();
        }
        return root.ToJsonString(Options);
    }

    /// <summary>Returns the common CLI outer object; batch items add only their index.</summary>
    public string ToCliJson(string command, long durationMs, int? index = null)
    {
        var root = new JsonObject
        {
            ["command"] = command,
            ["success"] = Success,
            ["durationMs"] = durationMs
        };
        if (index.HasValue)
            root["index"] = index.Value;
        if (HasData)
            root["data"] = Data?.DeepClone();
        if (Warnings is not null)
            root["warnings"] = Warnings.DeepClone();
        if (Error is not null)
        {
            var error = (JsonObject)Error.DeepClone();
            error.Remove("command");
            root["error"] = error;
        }
        if (PartialState is not null)
            root["partialState"] = PartialState.DeepClone();
        return root.ToJsonString(Options);
    }

    private static IEnumerable<JsonNode?> EnumerateAt(JsonNode? node, string[] path, int offset = 0)
    {
        if (node is JsonArray array)
        {
            foreach (var item in array)
                foreach (var value in EnumerateAt(item, path, offset))
                    yield return value;
        }
        else if (offset == path.Length)
            yield return node;
        else if (node is JsonObject obj && obj.TryGetPropertyValue(path[offset], out var child))
            foreach (var value in EnumerateAt(child, path, offset + 1))
                yield return value;
    }

    private static bool IsNegative(JsonObject? root) =>
        root?["success"] is JsonValue success && success.TryGetValue<bool>(out var ok) && !ok ||
        root?["isError"] is JsonValue error && error.TryGetValue<bool>(out var failed) && failed;

    private static string? ReadString(JsonObject? root, string name) =>
        root?[name] is JsonValue value && value.TryGetValue<string>(out var text) ? text : null;

    private static JsonArray? ReadWarnings(JsonObject? root)
    {
        var warnings = root?["warnings"] switch
        {
            JsonArray array => (JsonArray)array.DeepClone(),
            { } warning => new JsonArray(warning.DeepClone()),
            _ => null
        };
        if (root?["warning"] is { } singular)
        {
            warnings ??= [];
            warnings.Add(singular.DeepClone());
        }
        return warnings;
    }

    private static void Add(JsonObject root, string name, string? value)
    {
        if (value is not null)
            root[name] = value;
    }
}
