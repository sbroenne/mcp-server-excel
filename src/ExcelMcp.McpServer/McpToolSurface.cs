using System.Collections.Immutable;
using System.Reflection;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>
/// Derives the server's advertised tool and operation counts from the ACTUAL MCP tool
/// registration instead of hard-coded literals.
///
/// WHY THIS EXISTS
/// ---------------
/// The <c>--help</c> banner used to hard-code "Provides 22 tools with 195+ operations". The real
/// surface grew to 31 tools / 326 operations without anyone updating that string, so the binary
/// contradicted its own READMEs, SKILL.md files and its live <c>tools/list</c> response.
/// Reflecting over the registration makes that class of drift structurally impossible.
///
/// HOW THE NUMBERS ARE DERIVED
/// ---------------------------
/// Every MCP tool in this server is a single <c>[McpServerTool]</c> method on an
/// <c>[McpServerToolType]</c> class (hand-written or emitted by <c>McpToolGenerator</c>), and every
/// one of them dispatches on a required <c>action</c> parameter whose type is an action enum.
/// So:
///   tools      = number of [McpServerTool] methods
///   operations = sum of Enum.GetValues(actionEnumType).Length over those methods
///
/// This mirrors the discovery loop in <see cref="GeminiCompatibleToolRegistration"/>, which is what
/// actually registers the tools, so both walk the same surface.
/// </summary>
internal static class McpToolSurface
{
    /// <summary>Name of the dispatch parameter every tool exposes as an enum.</summary>
    private const string ActionParameterName = "action";

    /// <summary>One entry per registered MCP tool.</summary>
    /// <param name="Name">The tool name as advertised over the protocol (e.g. "range_format").</param>
    /// <param name="OperationCount">Number of values in the tool's <c>action</c> enum, or 0 if it has none.</param>
    internal sealed record ToolInfo(string Name, int OperationCount);

    private static readonly ImmutableArray<ToolInfo> DiscoveredTools = Discover(typeof(McpToolSurface).Assembly);

    /// <summary>All MCP tools registered by this server.</summary>
    public static ImmutableArray<ToolInfo> Tools => DiscoveredTools;

    /// <summary>Number of MCP tools this server registers.</summary>
    public static int ToolCount => DiscoveredTools.Length;

    /// <summary>Total number of operations (action enum values) across all registered tools.</summary>
    public static int OperationCount => DiscoveredTools.Sum(t => t.OperationCount);

    private static ImmutableArray<ToolInfo> Discover(Assembly toolAssembly)
    {
        const BindingFlags MethodFlags = BindingFlags.Public | BindingFlags.NonPublic |
                                         BindingFlags.Static | BindingFlags.Instance |
                                         BindingFlags.DeclaredOnly;

        var tools = ImmutableArray.CreateBuilder<ToolInfo>();

        foreach (var toolType in toolAssembly.GetTypes())
        {
            if (toolType.GetCustomAttribute<McpServerToolTypeAttribute>() is null)
            {
                continue;
            }

            foreach (var method in toolType.GetMethods(MethodFlags))
            {
                if (method.GetCustomAttribute<McpServerToolAttribute>() is not { } toolAttribute)
                {
                    continue;
                }

                tools.Add(new ToolInfo(
                    toolAttribute.Name ?? method.Name,
                    CountOperations(method)));
            }
        }

        return tools.ToImmutable();
    }

    /// <summary>
    /// Counts a tool's operations as the number of values in its <c>action</c> enum parameter.
    /// Returns 0 when a tool has no such parameter; <c>McpToolSurfaceTests</c> fails the build if
    /// that ever happens, so the total can never silently under-report.
    /// </summary>
    private static int CountOperations(MethodInfo method)
    {
        var actionParameter = Array.Find(
            method.GetParameters(),
            p => string.Equals(p.Name, ActionParameterName, StringComparison.Ordinal));

        if (actionParameter is null)
        {
            return 0;
        }

        var actionType = Nullable.GetUnderlyingType(actionParameter.ParameterType) ?? actionParameter.ParameterType;

        return actionType.IsEnum ? Enum.GetValues(actionType).Length : 0;
    }
}
