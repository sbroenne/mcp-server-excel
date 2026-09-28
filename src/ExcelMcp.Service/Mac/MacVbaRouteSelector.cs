using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed record MacVbaRoute(
    string? HelperAction,
    object? HelperArguments,
    string? UnavailableReason)
{
    public bool IsAvailable => HelperAction is not null;
}

internal static class MacVbaRouteSelector
{
    public static MacVbaRoute Select(
        string action,
        JsonObject arguments,
        MacVbaHelperCapabilities? capabilities,
        MacVbaPreflightResult preflight)
    {
        var helperAction = $"vba.{action}";
        if (action is not ("list" or "view" or "import" or "update" or "delete" or "run"))
        {
            return Unsupported("the action has no macOS VBA helper route");
        }
        if (action == "run")
        {
            if (preflight.MacroExecution != MacMacroExecutionAvailability.Available)
            {
                return Unsupported("macro execution is not enabled by the non-prompting preflight");
            }
        }
        else if (preflight.ProjectModelAccess != MacVbaProjectModelAccess.Enabled)
        {
            return Unsupported(
                "Trust access to the VBA project object model is not enabled");
        }
        if (capabilities is null)
        {
            return new(helperAction, CreateArguments(action, arguments), null);
        }
        if (action != "run" && !capabilities.ProjectReadable)
        {
            return Unsupported(
                "the helper cannot read the target workbook VBA project under current user-managed trust");
        }
        return capabilities.IsAvailable(helperAction)
            ? new(helperAction, CreateArguments(action, arguments), null)
            : Unsupported(
                $"helper method '{helperAction}' has not been proven or explicitly enabled");
    }

    private static object CreateArguments(string action, JsonObject arguments) =>
        action switch
        {
            "list" => new { },
            "view" or "delete" => new
            {
                moduleName = RequiredString(arguments, "moduleName")
            },
            "import" or "update" => new
            {
                moduleName = RequiredString(arguments, "moduleName"),
                source = RequiredString(arguments, "source")
            },
            "run" => new
            {
                procedureName = RequiredProcedure(arguments),
                parameters = ReadParameters(arguments)
            },
            _ => throw new ArgumentOutOfRangeException(nameof(action))
        };

    private static string RequiredProcedure(JsonObject arguments)
    {
        var value = RequiredString(arguments, "procedureName");
        var separator = value.IndexOf('.', StringComparison.Ordinal);
        if (separator <= 0
            || separator != value.LastIndexOf('.')
            || separator == value.Length - 1
            || !IsVbaIdentifier(value.AsSpan(0, separator), 31)
            || !IsVbaIdentifier(value.AsSpan(separator + 1), 255))
        {
            throw new ArgumentException(
                "procedureName must use exact 'Module.Procedure' form.");
        }
        return value;
    }

    private static bool IsVbaIdentifier(
        ReadOnlySpan<char> value,
        int maximumLength)
    {
        if (value.IsEmpty
            || value.Length > maximumLength
            || !IsAsciiLetter(value[0]))
        {
            return false;
        }
        for (var index = 1; index < value.Length; index++)
        {
            if (!IsAsciiLetter(value[index])
                && value[index] is not (>= '0' and <= '9') and not '_')
            {
                return false;
            }
        }
        return true;
    }

    private static bool IsAsciiLetter(char value) =>
        value is >= 'A' and <= 'Z' or >= 'a' and <= 'z';

    private static string[] ReadParameters(JsonObject arguments)
    {
        if (arguments["parameters"] is null)
        {
            return [];
        }
        if (arguments["parameters"] is not JsonArray values
            || values.Count > 30
            || values.Any(value => value is not JsonValue item
                || !item.TryGetValue<string>(out _)))
        {
            throw new ArgumentException(
                "parameters must contain no more than 30 strings.");
        }
        return values.Select(value => value!.GetValue<string>()).ToArray();
    }

    private static string RequiredString(JsonObject arguments, string propertyName) =>
        arguments[propertyName]?.GetValue<string>() is { } value
            && !string.IsNullOrWhiteSpace(value)
                ? value
                : throw new ArgumentException($"{propertyName} is required.");

    private static MacVbaRoute Unsupported(string reason) =>
        new(null, null, reason);
}
