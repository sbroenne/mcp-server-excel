using System.Reflection;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.McpServer.Tools;

namespace Sbroenne.ExcelMcp.McpServer;

internal static class McpActionContract
{
    private static readonly MethodInfo[] ManualMethods = new[] { typeof(ExcelFileTool), typeof(ExcelBatchTool) }
        .SelectMany(type => type.GetMethods(BindingFlags.Public | BindingFlags.Static))
        .Where(method => method.GetCustomAttribute<McpServerToolAttribute>() is not null)
        .ToArray();

    internal static (string Tool, string Action, bool RequiresSession)? GetOperation(string command)
    {
        if (ServiceRegistry.GetMcpOperation(command) is { } generated)
            return generated;
        foreach (var method in ManualMethods.Where(method => method.DeclaringType == typeof(ExcelFileTool)))
        {
            var tool = method.GetCustomAttribute<McpServerToolAttribute>()!.Name!;
            var actionType = method.GetParameters().Single(parameter => parameter.Name == "action").ParameterType;
            foreach (var value in Enum.GetValues(actionType))
            {
                var action = JsonSerializer.SerializeToElement(value).GetString()!;
                if (ServiceRegistry.GetMcpCommand(tool, action).Equals(command, StringComparison.OrdinalIgnoreCase))
                    return (tool, action, GetParameters(tool, action).Any(parameter =>
                        parameter.Name == "workbook_session_id" && parameter.Required));
            }
        }
        return null;
    }

    internal static (string Name, bool Required, bool AllowsEmpty, string? Alternative)[] GetParameters(string tool, string action)
    {
        var method = ManualMethods.SingleOrDefault(method =>
            method.GetCustomAttribute<McpServerToolAttribute>()!.Name == tool);
        if (method is null)
            return ServiceRegistry.GetMcpActionParameters(tool, action);
        return method.GetParameters()
            .SelectMany(parameter => parameter.GetCustomAttributes<McpActionParameterAttribute>()
                .Where(attribute => attribute.Action == action)
                .Select(attribute => (Name: parameter.Name!, Required: attribute.Required, AllowsEmpty: false, Alternative: (string?)null)))
            .Prepend(("action", true, false, (string?)null))
            .ToArray();
    }

    internal static void ValidateParameters(string tool, string action, IEnumerable<string> supplied)
    {
        var parameters = GetParameters(tool, action);
        var allowed = parameters.Select(parameter => parameter.Name).ToHashSet(StringComparer.Ordinal);
        var invalid = supplied.Where(name => !allowed.Contains(name)).Order(StringComparer.Ordinal).ToArray();
        if (invalid.Length > 0)
            throw new ArgumentException(
                $"Parameter(s) {string.Join(", ", invalid)} are not valid for {tool}.{action}. " +
                $"Valid action parameters: {string.Join(", ", allowed.Where(name => name is not ("action" or "workbook_session_id")).Order(StringComparer.Ordinal))}.");
    }
}
