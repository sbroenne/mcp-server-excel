using System.Reflection;
using System.Text.Json.Nodes;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>
/// Registers MCP tools with a schema generator configured for Google Gemini compatibility.
///
/// The MCP SDK's default schema generator (Microsoft.Extensions.AI <c>JsonSchemaExporter</c>)
/// expresses nullable .NET types using union types such as <c>"type": ["array","null"]</c>.
/// Some Gemini client adapters (including @ai-sdk/google 3.0.73) translate these into
/// <c>anyOf</c> while leaving <c>items</c> outside the array branch, causing issue #672.
///
/// <see cref="AIJsonSchemaTransformOptions.UseNullableKeyword"/> converts the union form into
/// the OpenAPI-style form (<c>type:"array"</c> + <c>nullable:true</c>), keeping arrays and
/// their items together. Nullable arrays are supported by Google's schema model; the
/// annotation is not a JSON Schema null alternative and may be ignored by adapters.
/// Untyped collection elements use explicit scalar alternatives, including a null type
/// for JSON Schema clients and adapters that translate it to OpenAPI nullability.
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
            // Avoid type arrays that some client adapters translate with misplaced items.
            UseNullableKeyword = true,
            TransformSchemaNode = (context, node) =>
            {
                if (node is JsonObject obj)
                {
                    // object cells generate {}, but typed alternatives describe the scalar values
                    // accepted at runtime without claiming that numbers and booleans become strings.
                    if (obj.Count == 0 && context.IsCollectionElementSchema)
                    {
                        obj["anyOf"] = new JsonArray(
                            new JsonObject { ["type"] = "string", ["nullable"] = true },
                            new JsonObject { ["type"] = "number" },
                            new JsonObject { ["type"] = "boolean" },
                            // nullable alone does not accept null in standard JSON Schema.
                            // A typed null branch is recognized by legacy Gemini adapters,
                            // unlike an untyped { enum: [null] } or { nullable: true } branch.
                            new JsonObject { ["type"] = "null" });
                    }
                }
                return node;
            }
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
