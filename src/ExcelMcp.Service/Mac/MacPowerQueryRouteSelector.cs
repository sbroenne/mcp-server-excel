using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal enum MacPowerQueryRouteKind
{
    SavedPackage,
    Helper,
    Unsupported
}

internal sealed record MacPowerQueryRoute(
    MacPowerQueryRouteKind Kind,
    string? HelperAction = null,
    JsonObject? HelperArguments = null,
    string? UnavailableReason = null);

internal static class MacPowerQueryRouteSelector
{
    public static MacPowerQueryRoute Select(
        string action,
        JsonObject arguments,
        IReadOnlySet<string> helperActions)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(action);
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentNullException.ThrowIfNull(helperActions);

        if (action is "list" or "view" or "get-load-config")
        {
            return SavedPackage();
        }

        if (action == "update")
        {
            var refresh = arguments["refresh"]?.GetValue<bool>() ?? true;
            if (!refresh)
            {
                return SavedPackage();
            }

            return Helper(
                "powerquery.update",
                new JsonObject
                {
                    ["name"] = RequiredString(arguments, "queryName"),
                    ["formula"] = RequiredString(arguments, "mCode"),
                    ["refresh"] = true
                },
                helperActions,
                "Power Query update with refresh requires the trusted helper refresh contract.");
        }

        if (action == "rename")
        {
            return Helper(
                "powerquery.rename",
                new JsonObject
                {
                    ["name"] = RequiredString(arguments, "oldName").Trim(),
                    ["newName"] = RequiredString(arguments, "newName").Trim()
                },
                helperActions,
                "Power Query rename requires the trusted helper.");
        }

        if (action == "delete")
        {
            return Helper(
                "powerquery.delete",
                new JsonObject
                {
                    ["name"] = RequiredString(arguments, "queryName"),
                    ["deleteConnection"] = true
                },
                helperActions,
                "Power Query delete requires the trusted helper cleanup contract.");
        }

        if (action == "create")
        {
            var destination = NormalizeDestination(
                arguments["loadDestination"]?.GetValue<string>() ?? "load-to-table");
            if (destination is "load-to-data-model" or "load-to-both")
            {
                return Unsupported(
                    "Power Query Data Model destinations require separate Mac engine evidence.");
            }

            var queryName = RequiredString(arguments, "queryName");
            var helperArguments = new JsonObject
            {
                ["name"] = queryName,
                ["formula"] = RequiredString(arguments, "mCode"),
                ["destination"] = destination,
                ["sheetName"] = null,
                ["cellAddress"] = null
            };
            if (destination == "load-to-table")
            {
                helperArguments["sheetName"] =
                    OptionalString(arguments, "targetSheet") ?? queryName;
                helperArguments["cellAddress"] =
                    OptionalString(arguments, "targetCellAddress") ?? "A1";
            }
            return Helper(
                "powerquery.create",
                helperArguments,
                helperActions,
                "Power Query create requires the trusted helper atomic create-and-load contract.");
        }

        if (action == "refresh")
        {
            return Helper(
                "powerquery.refresh",
                new JsonObject
                {
                    ["name"] = RequiredString(arguments, "queryName")
                },
                helperActions,
                "Power Query refresh requires a helper method that observes synchronous completion.");
        }

        if (action == "refresh-all")
        {
            return Helper(
                "powerquery.refresh-all",
                new JsonObject(),
                helperActions,
                "Power Query refresh-all requires per-query completion and error observation.");
        }

        if (action == "load-to")
        {
            var destination = NormalizeDestination(RequiredString(arguments, "loadDestination"));
            if (destination is "load-to-data-model" or "load-to-both")
            {
                return Unsupported(
                    "Power Query Data Model destinations require separate Mac engine evidence.");
            }

            var queryName = RequiredString(arguments, "queryName");
            if (destination == "connection-only"
                && OptionalString(arguments, "targetCellAddress") is not null)
            {
                throw new ArgumentException(
                    "targetCellAddress is only supported for worksheet loads.");
            }
            var helperArguments = new JsonObject
            {
                ["name"] = queryName,
                ["destination"] = destination,
                ["sheetName"] = null,
                ["cellAddress"] = null
            };
            if (destination == "load-to-table")
            {
                helperArguments["sheetName"] =
                    OptionalString(arguments, "targetSheet") ?? queryName;
                helperArguments["cellAddress"] =
                    OptionalString(arguments, "targetCellAddress") ?? "A1";
            }
            return Helper(
                "powerquery.load-to",
                helperArguments,
                helperActions,
                "Power Query load-to requires the trusted helper.");
        }

        if (action == "unload")
        {
            return Helper(
                "powerquery.unload",
                new JsonObject
                {
                    ["name"] = RequiredString(arguments, "queryName")
                },
                helperActions,
                "Power Query unload requires exact destination cleanup through the trusted helper.");
        }

        if (action == "evaluate")
        {
            return Helper(
                "powerquery.evaluate",
                new JsonObject
                {
                    ["formula"] = RequiredString(arguments, "mCode")
                },
                helperActions,
                "Power Query evaluate requires verified temporary-object cleanup through the trusted helper.");
        }

        var helperAction = $"powerquery.{action}";
        return Helper(
            helperAction,
            arguments.DeepClone().AsObject(),
            helperActions,
            $"Power Query {action} requires the trusted helper method '{helperAction}'.");
    }

    private static MacPowerQueryRoute SavedPackage() =>
        new(MacPowerQueryRouteKind.SavedPackage);

    private static MacPowerQueryRoute Helper(
        string action,
        JsonObject arguments,
        IReadOnlySet<string> helperActions,
        string unavailableReason) =>
        helperActions.Contains(action)
            ? new(MacPowerQueryRouteKind.Helper, action, arguments)
            : Unsupported(unavailableReason);

    private static MacPowerQueryRoute Unsupported(string reason) =>
        new(MacPowerQueryRouteKind.Unsupported, UnavailableReason: reason);

    private static string RequiredString(JsonObject arguments, string propertyName) =>
        arguments[propertyName]?.GetValue<string>() is { } value
            && !string.IsNullOrWhiteSpace(value)
                ? value
                : throw new ArgumentException($"{propertyName} is required.");

    private static string? OptionalString(JsonObject arguments, string propertyName) =>
        arguments[propertyName]?.GetValue<string>() is { } value
            && !string.IsNullOrWhiteSpace(value)
                ? value
                : null;

    private static string NormalizeDestination(string value) =>
        value.Trim().ToLowerInvariant() switch
        {
            "connection-only" => "connection-only",
            "load-to-table" or "worksheet" or "table" => "load-to-table",
            "load-to-data-model" or "data-model" or "datamodel" => "load-to-data-model",
            "load-to-both" or "both" => "load-to-both",
            _ => throw new ArgumentException($"Unknown Power Query load destination '{value}'.")
        };
}
