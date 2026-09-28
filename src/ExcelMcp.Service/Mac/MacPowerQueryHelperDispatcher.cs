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
        if (publicAction == "list")
        {
            result = MapListResult(workbookPath, helperObject);
        }
        else if (publicAction is "view" or "get-load-config")
        {
            var queryName = RequiredString(publicArguments, "queryName");
            result = MapReadResult(
                workbookPath,
                queryName,
                helperObject,
                includeFormula: publicAction == "view");
        }
        else if (publicAction == "rename")
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

    private static JsonObject MapListResult(
        string workbookPath,
        JsonObject helperResult)
    {
        if (helperResult["queries"] is not JsonArray queries)
        {
            throw new InvalidOperationException(
                "Helper list result is missing the queries array.");
        }

        var mappedQueries = new JsonArray();
        foreach (JsonNode? queryNode in queries)
        {
            if (queryNode is not JsonObject query)
            {
                throw new InvalidOperationException(
                    "Helper list result contains an invalid query.");
            }

            string loadMode = RequiredLoadMode(query);
            bool isConnectionOnly = RequiredBoolean(query, "isConnectionOnly");
            bool isLoadedToDataModel =
                RequiredBoolean(query, "isLoadedToDataModel");
            string? targetSheet = OptionalHelperString(query, "targetSheet");
            ValidateLoadMetadata(
                loadMode,
                isConnectionOnly,
                isLoadedToDataModel,
                hasConnection: null,
                targetSheet);
            var mapped = new JsonObject
            {
                ["name"] = RequiredHelperString(query, "name"),
                ["formulaPreview"] = RequiredHelperString(query, "formulaPreview"),
                ["characterCount"] = RequiredInt32(query, "characterCount"),
                ["loadMode"] = loadMode,
                ["isConnectionOnly"] = isConnectionOnly,
                ["isLoadedToDataModel"] = isLoadedToDataModel
            };
            if (targetSheet is not null)
            {
                mapped["targetSheet"] = targetSheet;
            }
            mappedQueries.Add(mapped);
        }

        var result = OperationResult(workbookPath);
        result["queries"] = mappedQueries;
        return result;
    }

    private static JsonObject MapReadResult(
        string workbookPath,
        string queryName,
        JsonObject helperResult,
        bool includeFormula)
    {
        string returnedName = RequiredHelperString(helperResult, "name");
        if (!string.Equals(queryName, returnedName, StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                "Helper view result does not identify the requested query.");
        }

        string loadMode = RequiredLoadMode(helperResult);
        bool hasConnection = RequiredBoolean(helperResult, "hasConnection");
        bool isConnectionOnly =
            RequiredBoolean(helperResult, "isConnectionOnly");
        bool isLoadedToDataModel =
            RequiredBoolean(helperResult, "isLoadedToDataModel");
        string? targetSheet =
            OptionalHelperString(helperResult, "targetSheet");
        ValidateLoadMetadata(
            loadMode,
            isConnectionOnly,
            isLoadedToDataModel,
            hasConnection,
            targetSheet);

        var result = OperationResult(workbookPath);
        result["queryName"] = queryName;
        result["loadMode"] = loadMode;
        result["hasConnection"] = hasConnection;
        result["isConnectionOnly"] = isConnectionOnly;
        result["isLoadedToDataModel"] = isLoadedToDataModel;
        if (targetSheet is not null)
        {
            result["targetSheet"] = targetSheet;
        }

        string formula = RequiredHelperString(helperResult, "formula");
        int characterCount = RequiredInt32(helperResult, "characterCount");
        if (formula.Length != characterCount)
        {
            throw new InvalidOperationException(
                "Helper view result formula length does not match characterCount.");
        }
        if (includeFormula)
        {
            result["mCode"] = formula;
            result["characterCount"] = characterCount;
        }

        return result;
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

    private static string? OptionalHelperString(
        JsonObject source,
        string propertyName)
    {
        if (source[propertyName] is null)
        {
            return null;
        }

        return RequiredHelperString(source, propertyName);
    }

    private static string RequiredString(JsonObject arguments, string propertyName) =>
        arguments[propertyName]?.GetValue<string>() is { } value
            ? value
            : throw new ArgumentException($"{propertyName} is required.");

    private static string RequiredHelperString(JsonObject value, string propertyName) =>
        value[propertyName] is JsonValue property
            && property.TryGetValue<string>(out var result)
            && !string.IsNullOrWhiteSpace(result)
                ? result
                : throw new InvalidOperationException(
                    $"Helper result property '{propertyName}' must be a non-empty string.");

    private static string RequiredLoadMode(JsonObject value)
    {
        string loadMode = RequiredHelperString(value, "loadMode");
        return loadMode is
                "connection-only" or
                "load-to-table" or
                "load-to-data-model" or
                "load-to-both"
            ? loadMode
            : throw new InvalidOperationException(
                "Helper result property 'loadMode' is invalid.");
    }

    private static void ValidateLoadMetadata(
        string loadMode,
        bool isConnectionOnly,
        bool isLoadedToDataModel,
        bool? hasConnection,
        string? targetSheet)
    {
        bool expectedConnectionOnly = loadMode == "connection-only";
        bool expectedDataModel =
            loadMode is "load-to-data-model" or "load-to-both";
        bool expectedConnection = !expectedConnectionOnly;
        bool expectedTargetSheet =
            loadMode is "load-to-table" or "load-to-both";
        if (isConnectionOnly != expectedConnectionOnly
            || isLoadedToDataModel != expectedDataModel
            || hasConnection is not null && hasConnection != expectedConnection
            || (targetSheet is not null) != expectedTargetSheet)
        {
            throw new InvalidOperationException(
                "Helper result contains inconsistent Power Query load metadata.");
        }
    }

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
