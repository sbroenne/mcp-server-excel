using System.Reflection;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>
/// Registers MCP tools with a schema generator configured for Google Gemini compatibility.
///
/// The MCP SDK's default schema generator (Microsoft.Extensions.AI <c>JsonSchemaExporter</c>)
/// expresses nullable .NET types using the JSON Schema 2020-12 union form
/// <c>"type": ["array","null"]</c>. Google's Gemini function-calling API only accepts a
/// subset of the OpenAPI 3.0.3 Schema object, which requires a SINGLE scalar <c>type</c> and
/// expresses nullability with a separate <c>nullable: true</c> field. Gemini rejects union
/// <c>type</c> arrays inside <c>items</c> with HTTP 400 (see issue #672).
///
/// <see cref="AIJsonSchemaTransformOptions.UseNullableKeyword"/> converts the union form into
/// the OpenAPI-3.0 form (<c>type:"array"</c> + <c>nullable:true</c>) that Gemini accepts, while
/// remaining valid for clients that speak full JSON Schema 2020-12.
///
/// The SDK's <c>WithToolsFromAssembly</c>/<c>WithTools</c> overloads do not expose a hook to set
/// <see cref="McpServerToolCreateOptions.SchemaCreateOptions"/>, so this method mirrors the SDK's
/// discovery loop and creates each tool with the Gemini-compatible schema options.
/// </summary>
internal static class GeminiCompatibleToolRegistration
{
    private static readonly AIJsonSchemaCreateOptions GeminiSchemaCreateOptions = new()
    {
        TransformOptions = new AIJsonSchemaTransformOptions
        {
            // Emit OpenAPI-3.0-style nullability (type:"array" + nullable:true) instead of
            // JSON-Schema-2020-12 union types (type:["array","null"]) that Gemini rejects.
            UseNullableKeyword = true
        }
    };

    /// <summary>
    /// Discovers all <c>[McpServerToolType]</c> classes in the given assembly and registers their
    /// <c>[McpServerTool]</c> methods using Gemini-compatible JSON schema generation.
    /// </summary>
    public static IMcpServerBuilder WithGeminiCompatibleToolsFromAssembly(
        this IMcpServerBuilder builder,
        Assembly? toolAssembly = null)
    {
        ArgumentNullException.ThrowIfNull(builder);

        toolAssembly ??= Assembly.GetCallingAssembly();

        var tools = new List<McpServerTool>();

        foreach (var toolType in toolAssembly.GetTypes())
        {
            if (toolType.GetCustomAttribute<McpServerToolTypeAttribute>() is null)
            {
                continue;
            }

            const BindingFlags methodFlags = BindingFlags.Public | BindingFlags.NonPublic |
                                             BindingFlags.Static | BindingFlags.Instance |
                                             BindingFlags.DeclaredOnly;

            foreach (var method in toolType.GetMethods(methodFlags))
            {
                if (method.GetCustomAttribute<McpServerToolAttribute>() is null)
                {
                    continue;
                }

                var options = new McpServerToolCreateOptions
                {
                    SchemaCreateOptions = GeminiSchemaCreateOptions
                };

                var tool = method.IsStatic
                    ? McpServerTool.Create(method, target: null, options)
                    : McpServerTool.Create(
                        method,
                        (RequestContext<CallToolRequestParams> request) =>
                            ActivatorUtilities.CreateInstance(
                                request.Services
                                    ?? throw new InvalidOperationException(
                                        "Request has no service provider to construct the tool instance."),
                                method.DeclaringType!),
                        options);

                tools.Add(tool);
            }
        }

        return builder.WithTools(tools);
    }
}
