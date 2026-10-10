using System.Reflection;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.McpServer.Tools;

namespace Sbroenne.ExcelMcp.McpServer;

internal static class McpActionContract
{
    private static readonly MethodInfo[] FileMethods = typeof(ExcelFileTool)
        .GetMethods(BindingFlags.Public | BindingFlags.Static)
        .Where(method => method.Name is nameof(ExcelFileTool.ExcelFile) or nameof(ExcelFileTool.ExcelFileRead))
        .ToArray();

    internal static (string Name, bool Required, bool AllowsEmpty, string? Alternative)[] GetParameters(string tool, string action)
    {
        if (tool is not ("file" or "file_read"))
            return ServiceRegistry.GetMcpActionParameters(tool, action);

        var methodName = tool == "file" ? nameof(ExcelFileTool.ExcelFile) : nameof(ExcelFileTool.ExcelFileRead);
        var method = FileMethods.Single(method => method.Name == methodName);
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
