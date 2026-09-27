using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed record MacVbaHelperRequest(
    int Version,
    string RequestId,
    string WorkbookPath,
    string Action,
    JsonElement Arguments);

internal sealed record MacVbaHelperError(
    string Category,
    string Code,
    string Message);

internal sealed record MacVbaHelperResponse(
    int Version,
    string RequestId,
    bool Success,
    JsonElement? Result,
    MacVbaHelperError? Error);

internal static class MacVbaHelperProtocol
{
    public const int Version = 1;
    public const int MaxPayloadBytes = 262_144;
    public const string HelperVersion = "1.0.0";

    private static readonly JsonSerializerOptions RequestJsonOptions =
        new(ServiceProtocol.JsonOptions)
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.Never
        };

    private static readonly HashSet<string> AllowedActions = new(StringComparer.Ordinal)
    {
        "helper.capabilities",
        "powerquery.list",
        "powerquery.view",
        "powerquery.create",
        "powerquery.update",
        "powerquery.rename",
        "powerquery.delete",
        "analysis.create-scenario",
        "analysis.show-scenario",
        "vba.list",
        "vba.view",
        "vba.import",
        "vba.update",
        "vba.delete"
    };

    private static readonly HashSet<string> RequestProperties = new(StringComparer.Ordinal)
    {
        "version",
        "requestId",
        "workbookPath",
        "action",
        "arguments"
    };

    private static readonly HashSet<string> ResponseProperties = new(StringComparer.Ordinal)
    {
        "version",
        "requestId",
        "success",
        "result",
        "error"
    };

    public static string CreateRequest(
        string requestId,
        string workbookPath,
        string action,
        object arguments)
    {
        ValidateRequestId(requestId);
        ArgumentException.ThrowIfNullOrWhiteSpace(workbookPath);
        if (!Path.IsPathFullyQualified(workbookPath))
        {
            throw new ArgumentException("workbookPath must be an absolute path.", nameof(workbookPath));
        }
        ValidateAction(action);
        ArgumentNullException.ThrowIfNull(arguments);

        var json = JsonSerializer.Serialize(
            new
            {
                version = Version,
                requestId,
                workbookPath,
                action,
                arguments
            },
            RequestJsonOptions);
        EnsurePayloadSize(json, "Helper request");
        return json;
    }

    public static MacVbaHelperRequest ParseRequest(string json)
    {
        EnsurePayloadSize(json, "Helper request");
        using var document = JsonDocument.Parse(json, new JsonDocumentOptions
        {
            AllowTrailingCommas = false,
            CommentHandling = JsonCommentHandling.Disallow,
            MaxDepth = 32
        });
        var root = document.RootElement;
        ValidateObjectProperties(root, RequestProperties, "helper request");

        var version = root.GetProperty("version").GetInt32();
        if (version != Version)
        {
            throw new ArgumentException($"Unsupported helper protocol version '{version}'.");
        }
        var requestId = root.GetProperty("requestId").GetString() ?? "";
        ValidateRequestId(requestId);
        var workbookPath = root.GetProperty("workbookPath").GetString() ?? "";
        if (!Path.IsPathFullyQualified(workbookPath))
        {
            throw new ArgumentException("workbookPath must be an absolute path.");
        }
        var action = root.GetProperty("action").GetString() ?? "";
        ValidateAction(action);
        var arguments = root.GetProperty("arguments");
        if (arguments.ValueKind != JsonValueKind.Object)
        {
            throw new ArgumentException("Helper request arguments must be a JSON object.");
        }

        return new MacVbaHelperRequest(
            version,
            requestId,
            workbookPath,
            action,
            arguments.Clone());
    }

    public static MacVbaHelperResponse ParseResponse(string json, string expectedRequestId)
    {
        ValidateRequestId(expectedRequestId);
        EnsurePayloadSize(json, "Helper response");
        using var document = JsonDocument.Parse(json, new JsonDocumentOptions
        {
            AllowTrailingCommas = false,
            CommentHandling = JsonCommentHandling.Disallow,
            MaxDepth = 32
        });
        var root = document.RootElement;
        ValidateObjectProperties(root, ResponseProperties, "helper response");

        var version = root.GetProperty("version").GetInt32();
        if (version != Version)
        {
            throw new InvalidOperationException($"Helper returned unsupported protocol version '{version}'.");
        }
        var requestId = root.GetProperty("requestId").GetString() ?? "";
        if (!string.Equals(requestId, expectedRequestId, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Helper response correlation did not match the request.");
        }

        var success = root.GetProperty("success").GetBoolean();
        var resultElement = root.GetProperty("result");
        var errorElement = root.GetProperty("error");
        if (success)
        {
            if (errorElement.ValueKind != JsonValueKind.Null)
            {
                throw new InvalidOperationException("A successful helper response must have a null error.");
            }
            return new MacVbaHelperResponse(
                version,
                requestId,
                true,
                resultElement.ValueKind == JsonValueKind.Null ? null : resultElement.Clone(),
                null);
        }

        if (resultElement.ValueKind != JsonValueKind.Null
            || errorElement.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidOperationException(
                "A failed helper response must have a null result and structured error.");
        }
        ValidateObjectProperties(
            errorElement,
            new HashSet<string>(["category", "code", "message"], StringComparer.Ordinal),
            "helper error");
        var error = new MacVbaHelperError(
            errorElement.GetProperty("category").GetString() ?? "ComInterop",
            errorElement.GetProperty("code").GetString() ?? "unknown",
            errorElement.GetProperty("message").GetString() ?? "Helper operation failed.");
        return new MacVbaHelperResponse(version, requestId, false, null, error);
    }

    private static void ValidateRequestId(string requestId)
    {
        if (requestId.Length != 32
            || requestId.Any(character => character is not (>= '0' and <= '9')
                and not (>= 'a' and <= 'f')))
        {
            throw new ArgumentException(
                "requestId must be 32 lowercase hexadecimal characters.",
                nameof(requestId));
        }
    }

    private static void ValidateAction(string action)
    {
        if (!AllowedActions.Contains(action))
        {
            throw new ArgumentException(
                $"Helper action '{action}' is not allowed.",
                nameof(action));
        }
    }

    private static void EnsurePayloadSize(string value, string payloadName)
    {
        ArgumentNullException.ThrowIfNull(value);
        var byteCount = Encoding.UTF8.GetByteCount(value);
        if (byteCount > MaxPayloadBytes)
        {
            throw new ArgumentException(
                $"{payloadName} is {byteCount} UTF-8 bytes; the initial transport limit is " +
                $"{MaxPayloadBytes} bytes. Reduce inline source or query text and retry.");
        }
    }

    private static void ValidateObjectProperties(
        JsonElement element,
        HashSet<string> allowedProperties,
        string objectName)
    {
        if (element.ValueKind != JsonValueKind.Object)
        {
            throw new ArgumentException($"{objectName} must be a JSON object.");
        }

        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var property in element.EnumerateObject())
        {
            if (!seen.Add(property.Name))
            {
                throw new ArgumentException(
                    $"{objectName} contains duplicate property '{property.Name}'.");
            }
            if (!allowedProperties.Contains(property.Name))
            {
                throw new ArgumentException(
                    $"{objectName} contains unknown property '{property.Name}'.");
            }
        }
        if (seen.Count != allowedProperties.Count)
        {
            throw new ArgumentException($"{objectName} is missing a required property.");
        }
    }
}
