using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.McpServer.Tools;

namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>
/// Reports missing session identity before the SDK's parameter binder hides the error.
/// </summary>
internal static class SessionIdentityFilter
{
    internal const string ErrorMessage =
        "A non-empty string session_id is required in the tools/call arguments object. " +
        "Use the ID returned by file open/create or the matching entry from file list on this MCP server. " +
        "If you supplied it, check that the client or bridge forwards session_id unchanged.";

    internal static McpRequestHandler<CallToolRequestParams, CallToolResult> Wrap(
        McpRequestHandler<CallToolRequestParams, CallToolResult> next) =>
        (request, cancellationToken) =>
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (request.MatchedPrimitive is McpServerTool tool &&
                request.Params?.Arguments is { } arguments &&
                arguments.TryGetValue("action", out var action) &&
                IsDeclaredAction(tool.ProtocolTool.InputSchema, action) &&
                RequiresSessionIdentity(tool.ProtocolTool, action) &&
                (!arguments.TryGetValue("session_id", out var sessionId) ||
                 sessionId.ValueKind != JsonValueKind.String ||
                 string.IsNullOrWhiteSpace(sessionId.GetString())))
            {
                // Validate before string binding, including file.close where the
                // tool-level schema makes identity optional for the other actions.
                var error = ExcelToolsBase.SerializeToolError(
                    "tools/call", null, new ArgumentException(ErrorMessage));
                return ValueTask.FromResult(new CallToolResult
                {
                    IsError = true,
                    Content = [new TextContentBlock { Text = error }]
                });
            }

            return next(request, cancellationToken);
        };

    private static bool RequiresSessionIdentity(Tool tool, JsonElement action) =>
        (tool.InputSchema.TryGetProperty("required", out var required) &&
         required.EnumerateArray().Any(property => property.GetString() == "session_id")) ||
        (tool.Name == "file" &&
         string.Equals(action.GetString(), "close", StringComparison.OrdinalIgnoreCase));

    private static bool IsDeclaredAction(JsonElement schema, JsonElement action)
    {
        // Leave missing/invalid actions to the SDK, including its existing error ordering.
        return action.ValueKind == JsonValueKind.String &&
            schema.GetProperty("properties").GetProperty("action").TryGetProperty("enum", out var actions) &&
            actions.EnumerateArray().Any(value =>
                string.Equals(value.GetString(), action.GetString(), StringComparison.OrdinalIgnoreCase));
    }
}
