using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacPythonInExcelArguments
{
    public static void Prepare(string action, JsonObject arguments, TimeSpan operationTimeout)
    {
        RequireString(arguments, "sheetName");
        RequireString(arguments, "rangeAddress");

        if (action == "set-formula")
        {
            var code = arguments["code"]?.GetValue<string>();
            ValidateCode(code);
            arguments["returnType"] ??= 0;
            return;
        }

        if (action != "get-result")
        {
            throw new ArgumentException($"Unknown Python in Excel action: {action}", nameof(action));
        }

        var maxWaitSeconds = arguments["maxWaitSeconds"]?.GetValue<int>() ?? 30;
        ValidateMaxWaitSeconds(maxWaitSeconds, operationTimeout);
        arguments["maxWaitSeconds"] = maxWaitSeconds;
        arguments["operationTimeoutSeconds"] = operationTimeout.TotalSeconds;
    }

    private static void ValidateCode(string? code)
    {
        if (string.IsNullOrWhiteSpace(code))
        {
            throw new ArgumentException("Python code must not be empty.", nameof(code));
        }
    }

    private static void ValidateMaxWaitSeconds(int maxWaitSeconds, TimeSpan operationTimeout)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(maxWaitSeconds, 1);
        if (TimeSpan.FromSeconds(maxWaitSeconds) >= operationTimeout)
        {
            throw new ArgumentOutOfRangeException(
                nameof(maxWaitSeconds),
                maxWaitSeconds,
                $"maxWaitSeconds must be less than the session operation timeout of {operationTimeout.TotalSeconds:0.###} seconds. " +
                "Reopen the session with a larger --timeout, or use a shorter wait and call get-result again.");
        }
    }

    private static string RequireString(JsonObject arguments, string propertyName)
    {
        var value = arguments[propertyName]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException($"{propertyName} is required.", propertyName);
        }

        return value;
    }
}
