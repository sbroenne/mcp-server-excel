using System.Text.Json;
using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacVbaHelperDispatcher(MacPowerQueryHelperDispatch dispatch)
{
    public async Task<JsonElement> DispatchAsync(
        MacVbaRoute route,
        string workbookPath,
        TimeSpan timeout,
        string publicAction)
    {
        if (!route.IsAvailable
            || route.HelperAction is null
            || route.HelperArguments is null)
        {
            throw new ArgumentException("A complete helper route is required.", nameof(route));
        }

        var helperResult = await dispatch(
            workbookPath,
            route.HelperAction,
            route.HelperArguments,
            timeout);
        var helperObject = JsonNode.Parse(helperResult.GetRawText())?.AsObject()
            ?? throw new InvalidOperationException(
                $"Helper action '{route.HelperAction}' returned a non-object result.");

        JsonObject result = publicAction switch
        {
            "list" => MapList(helperObject, workbookPath),
            "view" => MapView(helperObject, workbookPath),
            "import" or "update" or "delete" or "run" =>
                OperationResult(helperObject, workbookPath),
            _ => throw new ArgumentOutOfRangeException(nameof(publicAction))
        };
        return JsonSerializer.SerializeToElement(result, ServiceProtocol.JsonOptions);
    }

    private static JsonObject MapList(JsonObject helperResult, string workbookPath)
    {
        if (helperResult["modules"] is not JsonArray modules)
        {
            throw new InvalidOperationException(
                "Helper VBA list result is missing modules.");
        }
        var scripts = new JsonArray();
        foreach (var module in modules)
        {
            if (module is not JsonObject value)
            {
                throw new InvalidOperationException(
                    "Helper VBA list result contains an invalid module.");
            }
            scripts.Add(new JsonObject
            {
                ["name"] = RequiredString(value, "name"),
                ["type"] = ModuleType(RequiredInt32(value, "type")),
                ["lineCount"] = RequiredNonNegativeInt32(value, "lineCount"),
                ["procedures"] = RequiredStringArray(value, "procedures")
            });
        }
        var result = OperationResult(new JsonObject(), workbookPath);
        result["scripts"] = scripts;
        return result;
    }

    private static JsonObject MapView(JsonObject helperResult, string workbookPath)
    {
        var result = OperationResult(new JsonObject(), workbookPath);
        result["moduleName"] = RequiredString(helperResult, "moduleName");
        result["moduleType"] = ModuleType(RequiredInt32(helperResult, "moduleType"));
        result["code"] = RequiredStringAllowEmpty(helperResult, "source");
        result["lineCount"] = RequiredNonNegativeInt32(helperResult, "lineCount");
        result["procedures"] = RequiredStringArray(helperResult, "procedures");
        return result;
    }

    private static JsonObject OperationResult(
        JsonObject helperResult,
        string workbookPath)
    {
        if (helperResult.Count != 0)
        {
            throw new InvalidOperationException(
                "Helper VBA mutation result must be an empty object.");
        }
        return new()
        {
            ["success"] = true,
            ["filePath"] = workbookPath
        };
    }

    private static string ModuleType(int value) =>
        value switch
        {
            1 => "Module",
            2 => "Class",
            3 => "Form",
            100 => "Document",
            _ => $"Type{value}"
        };

    private static JsonArray RequiredStringArray(
        JsonObject value,
        string propertyName)
    {
        if (value[propertyName] is not JsonArray array
            || array.Any(item => item is not JsonValue jsonValue
                || !jsonValue.TryGetValue<string>(out _)))
        {
            throw new InvalidOperationException(
                $"Helper result property '{propertyName}' must be a string array.");
        }
        return (JsonArray)array.DeepClone();
    }

    private static string RequiredString(JsonObject value, string propertyName) =>
        RequiredStringAllowEmpty(value, propertyName) is { Length: > 0 } result
            ? result
            : throw new InvalidOperationException(
                $"Helper result property '{propertyName}' must be a non-empty string.");

    private static string RequiredStringAllowEmpty(
        JsonObject value,
        string propertyName) =>
        value[propertyName] is JsonValue property
        && property.TryGetValue<string>(out var result)
            ? result
            : throw new InvalidOperationException(
                $"Helper result property '{propertyName}' must be a string.");

    private static int RequiredInt32(JsonObject value, string propertyName) =>
        value[propertyName] is JsonValue property
        && property.TryGetValue<int>(out var result)
            ? result
            : throw new InvalidOperationException(
                $"Helper result property '{propertyName}' must be an integer.");

    private static int RequiredNonNegativeInt32(
        JsonObject value,
        string propertyName)
    {
        var result = RequiredInt32(value, propertyName);
        return result >= 0
            ? result
            : throw new InvalidOperationException(
                $"Helper result property '{propertyName}' must be non-negative.");
    }
}
