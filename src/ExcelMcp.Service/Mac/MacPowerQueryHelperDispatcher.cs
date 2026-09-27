using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal delegate Task<JsonElement> MacPowerQueryHelperDispatch(
    string workbookPath,
    string action,
    object arguments,
    TimeSpan timeout);

internal sealed class MacPowerQueryHelperDispatcher(
    MacPowerQueryHelperDispatch dispatch)
{
    public async Task<JsonElement> DispatchAsync(
        MacPowerQueryRoute route,
        string workbookPath,
        TimeSpan timeout,
        string publicAction,
        JsonObject publicArguments)
    {
        if (route.Kind != MacPowerQueryRouteKind.Helper
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
        JsonObject result;
        if (publicAction == "rename")
        {
            var oldName = RequiredString(publicArguments, "oldName");
            var newName = RequiredString(publicArguments, "newName");
            result = OperationResult(workbookPath);
            result["objectType"] = "power-query";
            result["oldName"] = oldName;
            result["newName"] = newName;
            result["normalizedOldName"] = oldName.Trim();
            result["normalizedNewName"] = newName.Trim();
        }
        else if (publicAction == "refresh")
        {
            ValidateRefreshResult(helperObject, publicArguments);
            result = OperationResult(workbookPath);
            CopyRequired(helperObject, result, "queryName");
            CopyRequired(helperObject, result, "hasErrors");
            CopyRequired(helperObject, result, "errorMessages");
            CopyRequired(helperObject, result, "refreshTime");
            CopyRequired(helperObject, result, "isConnectionOnly");
            CopyOptional(helperObject, result, "loadedToSheet");
        }
        else if (publicAction == "evaluate")
        {
            ValidateEvaluateResult(helperObject);
            result = OperationResult(workbookPath);
            result["mCode"] = RequiredString(publicArguments, "mCode");
            CopyRequired(helperObject, result, "columns");
            CopyRequired(helperObject, result, "rows");
            CopyRequired(helperObject, result, "rowCount");
            CopyRequired(helperObject, result, "columnCount");
        }
        else
        {
            result = OperationResult(workbookPath);
            if (publicAction == "unload")
            {
                result["action"] = "unload";
            }
        }

        return JsonSerializer.SerializeToElement(result, ServiceProtocol.JsonOptions);
    }

    private static JsonObject OperationResult(string workbookPath) =>
        new()
        {
            ["success"] = true,
            ["filePath"] = workbookPath
        };

    private static void ValidateRefreshResult(
        JsonObject helperResult,
        JsonObject publicArguments)
    {
        var requestedName = RequiredString(publicArguments, "queryName");
        var returnedName = RequiredString(helperResult, "queryName");
        if (!string.Equals(requestedName, returnedName, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "Helper refresh result does not identify the requested query.");
        }
        if (RequiredBoolean(helperResult, "hasErrors"))
        {
            throw new InvalidOperationException(
                "Helper refresh reported query errors in a successful response.");
        }
        if (helperResult["errorMessages"] is not JsonArray errorMessages
            || errorMessages.Count != 0)
        {
            throw new InvalidOperationException(
                "A successful helper refresh must return an empty errorMessages array.");
        }
        var refreshTime = RequiredString(helperResult, "refreshTime");
        if (!DateTimeOffset.TryParse(
                refreshTime,
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out _))
        {
            throw new InvalidOperationException(
                "Helper refresh result contains an invalid refreshTime.");
        }
        _ = RequiredBoolean(helperResult, "isConnectionOnly");
    }

    private static void ValidateEvaluateResult(JsonObject helperResult)
    {
        if (helperResult["columns"] is not JsonArray columns
            || columns.Any(column => column is not JsonValue value
                || !value.TryGetValue<string>(out _)))
        {
            throw new InvalidOperationException(
                "Helper evaluate result contains invalid columns.");
        }
        if (helperResult["rows"] is not JsonArray rows
            || rows.Any(row => row is not JsonArray cells
                || cells.Count != columns.Count))
        {
            throw new InvalidOperationException(
                "Helper evaluate result contains invalid rows.");
        }
        if (RequiredInt32(helperResult, "rowCount") != rows.Count
            || RequiredInt32(helperResult, "columnCount") != columns.Count)
        {
            throw new InvalidOperationException(
                "Helper evaluate result dimensions do not match its data.");
        }
    }

    private static void CopyRequired(
        JsonObject source,
        JsonObject destination,
        string propertyName)
    {
        destination[propertyName] = source[propertyName]?.DeepClone()
            ?? throw new InvalidOperationException(
                $"Helper result is missing required property '{propertyName}'.");
    }

    private static void CopyOptional(
        JsonObject source,
        JsonObject destination,
        string propertyName)
    {
        if (source[propertyName] is { } value)
        {
            destination[propertyName] = value.DeepClone();
        }
    }

    private static string RequiredString(JsonObject arguments, string propertyName) =>
        arguments[propertyName]?.GetValue<string>() is { } value
            ? value
            : throw new ArgumentException($"{propertyName} is required.");

    private static bool RequiredBoolean(JsonObject value, string propertyName) =>
        value[propertyName] is JsonValue property
            && property.TryGetValue<bool>(out var result)
                ? result
                : throw new InvalidOperationException(
                    $"Helper result property '{propertyName}' must be a Boolean.");

    private static int RequiredInt32(JsonObject value, string propertyName) =>
        value[propertyName] is JsonValue property
            && property.TryGetValue<int>(out var result)
            && result >= 0
                ? result
                : throw new InvalidOperationException(
                    $"Helper result property '{propertyName}' must be a non-negative integer.");
}
