using System.Globalization;
using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacNamedRangeArguments
{
    public static void Prepare(string action, JsonObject arguments)
    {
        if (action is not ("list" or "create" or "read" or "write" or "update" or "delete"))
        {
            throw new ArgumentException($"Unknown named range action '{action}'.", nameof(action));
        }
        if (action == "list")
        {
            return;
        }

        var name = RequiredText(arguments, "name");
        if (action is "create" or "update")
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(name);
            if (name.Length > 255)
            {
                throw new ArgumentException(
                    $"Named range name exceeds Excel's 255-character limit (current length: {name.Length}).",
                    nameof(arguments));
            }
            arguments["reference"] = "=" + RequiredText(arguments, "reference").TrimStart('=');
        }
        else if (action == "write")
        {
            var value = RequiredText(arguments, "value");
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            {
                if (!double.IsFinite(number))
                {
                    throw new ArgumentException("Named range numeric values must be finite.", nameof(arguments));
                }
                arguments["parsedValue"] = number;
            }
            else if (bool.TryParse(value, out var boolean))
            {
                arguments["parsedValue"] = boolean;
            }
            else
            {
                arguments["parsedValue"] = value;
            }
        }
    }

    public static void PrepareRangeBinding(string action, JsonObject arguments, bool enabled)
    {
        // Internal bindings must come from validated public arguments, not caller-supplied keys.
        arguments.Remove("namedRangeName");
        if (action is not ("get-values" or "set-values")
            || arguments["sheetName"]?.GetValue<string>() != string.Empty)
        {
            return;
        }
        if (!enabled)
        {
            throw new PlatformNotSupportedException(
                "Native named-range bulk addresses require the corresponding named-range read or write capability.");
        }
        var name = RequiredText(arguments, "rangeAddress");
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        arguments["namedRangeName"] = name;
    }

    private static string RequiredText(JsonObject arguments, string name)
    {
        if (arguments[name] is JsonValue value && value.TryGetValue<string>(out var text))
        {
            return text;
        }
        throw new ArgumentException($"'{name}' is required and must be a string.", nameof(arguments));
    }
}
