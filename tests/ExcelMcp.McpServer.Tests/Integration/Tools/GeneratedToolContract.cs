using System.Reflection;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

internal static class GeneratedToolContract
{
    internal static ParameterInfo GetParameter(string toolName, string parameterName)
    {
        var method = typeof(Program).Assembly.GetTypes()
            .Where(type => type.GetCustomAttribute<McpServerToolTypeAttribute>() != null)
            .SelectMany(type => type.GetMethods(BindingFlags.Public | BindingFlags.Static))
            .Single(method => method.GetCustomAttribute<McpServerToolAttribute>()?.Name == toolName);

        return method.GetParameters().Single(parameter => parameter.Name == parameterName);
    }
}
