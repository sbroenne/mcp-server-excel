using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.McpServer.Tools;

namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>
/// Reports missing session identity before the SDK's parameter binder hides the error.
/// </summary>
internal static class SessionIdentityFilter
{
    private static readonly Dictionary<string, HashSet<string>> ConditionalSessionActions = new(StringComparer.Ordinal)
    {
        ["file"] = new(StringComparer.OrdinalIgnoreCase) { "close" },
        ["worksheet"] = new(StringComparer.OrdinalIgnoreCase)
        {
            ServiceRegistry.Sheet.ListAction,
            ServiceRegistry.Sheet.CreateAction,
            ServiceRegistry.Sheet.RenameAction,
            ServiceRegistry.Sheet.DeleteAction,
            ServiceRegistry.Sheet.MoveAction,
            ServiceRegistry.Sheet.CopyAction
        }
    };

    internal const string ErrorMessage =
        "A non-empty string session_id is required in the tools/call arguments object. " +
        "Use the ID returned by file open/create or the matching entry from file list on this MCP server. " +
        "If you supplied it, check that the client or bridge forwards session_id unchanged.";

    internal const string AmbiguousErrorMessage =
        "The tools/call arguments object contains ambiguous session identity. " +
        "Use a non-empty string session_id. If the compatibility alias sessionId is also present, " +
        "it must be a non-empty string with the identical value.";

    internal static McpRequestHandler<CallToolRequestParams, CallToolResult> Wrap(
        McpRequestHandler<CallToolRequestParams, CallToolResult> next) =>
        (request, cancellationToken) =>
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (request.MatchedPrimitive is McpServerTool tool &&
                request.Params?.Arguments is { } arguments &&
                arguments.TryGetValue("action", out var action) &&
                IsDeclaredAction(tool.ProtocolTool.InputSchema, action) &&
                RequiresSessionIdentity(tool.ProtocolTool, action))
            {
                var validationError = NormalizeSessionIdentity(
                    arguments,
                    tool.ProtocolTool.Name,
                    action.GetString()!,
                    (toolName, actionName) => ObserveAlias(request, toolName, actionName));
                if (validationError is not null)
                {
                    // Validate before string binding, including action-specific requirements
                    // hidden by an optional tool-level session_id.
                    var error = ExcelToolsBase.SerializeToolError(
                        "tools/call", null, new ArgumentException(validationError));
                    return ValueTask.FromResult(new CallToolResult
                    {
                        IsError = true,
                        Content = [new TextContentBlock { Text = error }]
                    });
                }
            }

            return next(request, cancellationToken);
        };

    internal static string? NormalizeSessionIdentity(
        IDictionary<string, JsonElement> arguments,
        string toolName,
        string action,
        Action<string, string> aliasObserved)
    {
        var hasCanonical = arguments.TryGetValue("session_id", out var canonical);
        var hasAlias = arguments.TryGetValue("sessionId", out var alias);

        if (hasAlias)
        {
            aliasObserved(toolName, action);
        }

        if (hasCanonical && hasAlias)
        {
            return IsNonBlankString(canonical) &&
                   IsNonBlankString(alias) &&
                   string.Equals(canonical.GetString(), alias.GetString(), StringComparison.Ordinal)
                ? null
                : AmbiguousErrorMessage;
        }

        if (hasCanonical)
        {
            return IsNonBlankString(canonical) ? null : ErrorMessage;
        }

        if (!hasAlias || !IsNonBlankString(alias))
        {
            return ErrorMessage;
        }

        arguments["session_id"] = alias;
        return null;
    }

    private static bool IsNonBlankString(JsonElement value) =>
        value.ValueKind == JsonValueKind.String &&
        !string.IsNullOrWhiteSpace(value.GetString());

    private static void ObserveAlias(
        RequestContext<CallToolRequestParams> request,
        string toolName,
        string action)
    {
        var logger = request.Services?
            .GetService<ILoggerFactory>()?
            .CreateLogger(typeof(SessionIdentityFilter).FullName!);
        WriteAliasWarning(logger, toolName, action);
        ExcelMcpTelemetry.TrackSessionIdAliasObserved(toolName, action);
    }

    internal static void WriteAliasWarning(ILogger? logger, string toolName, string action)
    {
        try
        {
            logger?.LogWarning(
                "Compatibility sessionId alias observed for MCP tool {Tool}/{Action}; use session_id.",
                toolName,
                action);
        }
        catch (Exception)
        {
            // Compatibility diagnostics must never affect tool execution.
        }
    }

    private static bool RequiresSessionIdentity(Tool tool, JsonElement action) =>
        (tool.InputSchema.TryGetProperty("required", out var required) &&
         required.EnumerateArray().Any(property => property.GetString() == "session_id")) ||
        (ConditionalSessionActions.TryGetValue(tool.Name, out var actions) &&
         actions.Contains(action.GetString()!));

    private static bool IsDeclaredAction(JsonElement schema, JsonElement action)
    {
        // Leave missing/invalid actions to the SDK, including its existing error ordering.
        return action.ValueKind == JsonValueKind.String &&
            schema.GetProperty("properties").GetProperty("action").TryGetProperty("enum", out var actions) &&
            actions.EnumerateArray().Any(value =>
                string.Equals(value.GetString(), action.GetString(), StringComparison.OrdinalIgnoreCase));
    }
}
