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
        var result = JsonNode.Parse(helperResult.GetRawText())?.AsObject()
            ?? throw new InvalidOperationException(
                $"Helper action '{route.HelperAction}' returned a non-object result.");

        result["success"] = true;
        result["filePath"] = workbookPath;
        if (publicAction == "rename")
        {
            var oldName = RequiredString(publicArguments, "oldName");
            var newName = RequiredString(publicArguments, "newName");
            result["objectType"] = "power-query";
            result["oldName"] = oldName;
            result["newName"] = newName;
            result["normalizedOldName"] = oldName.Trim();
            result["normalizedNewName"] = newName.Trim();
        }

        return JsonSerializer.SerializeToElement(result, ServiceProtocol.JsonOptions);
    }

    private static string RequiredString(JsonObject arguments, string propertyName) =>
        arguments[propertyName]?.GetValue<string>() is { } value
            ? value
            : throw new ArgumentException($"{propertyName} is required.");
}
