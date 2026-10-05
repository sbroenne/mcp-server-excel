using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.McpServer.Tools;

namespace Sbroenne.ExcelMcp.McpServer;

internal static class ToolArgumentFilter
{
    internal static McpRequestHandler<CallToolRequestParams, CallToolResult> Wrap(
        McpRequestHandler<CallToolRequestParams, CallToolResult> next) =>
        (request, cancellationToken) =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (request.MatchedPrimitive is not McpServerTool tool)
                return next(request, cancellationToken);

            try
            {
                var arguments = request.Params?.Arguments
                    ?? throw new ArgumentException("The action argument is required.");
                var properties = tool.ProtocolTool.InputSchema.GetProperty("properties");
                if (!arguments.TryGetValue("action", out var action) || action.ValueKind != JsonValueKind.String)
                    throw new ArgumentException("The action argument must be a string naming an available action.");
                var canonicalAction = properties.GetProperty("action").GetProperty("enum").EnumerateArray()
                    .Select(value => value.GetString()!)
                    .FirstOrDefault(value => string.Equals(value, action.GetString(), StringComparison.OrdinalIgnoreCase))
                    ?? throw new ArgumentException("Unknown action. Use an action from this tool's schema.");

                foreach (var (name, value) in arguments)
                {
                    if (!properties.TryGetProperty(name, out var schema))
                        throw new ArgumentException($"Unknown parameter '{name}'.");
                    ValidateValueKind(name, value, schema);
                    if (name == "workbook_session_id" && value.ValueKind == JsonValueKind.String
                        && string.IsNullOrWhiteSpace(value.GetString()))
                        throw new ArgumentException("workbook_session_id must be a non-empty string.");
                }

                foreach (var required in tool.ProtocolTool.InputSchema.GetProperty("required").EnumerateArray())
                {
                    var name = required.GetString()!;
                    if (!arguments.ContainsKey(name))
                        throw new ArgumentException($"Parameter '{name}' is required.");
                }

                if (tool.ProtocolTool.Name is "file" or "file_read")
                {
                    ExcelFileTool.ValidateActionParameters(tool.ProtocolTool.Name, canonicalAction, arguments.Keys);
                }
                else
                {
                    var names = arguments.Keys.Where(name => name is not ("action" or "workbook_session_id")).ToArray();
                    try
                    {
                        ServiceRegistry.ValidateMcpActionParameters(
                            tool.ProtocolTool.Name is "worksheet" or "worksheet_read"
                                ? ServiceRegistry.Sheet.McpToolName
                                : tool.ProtocolTool.Name,
                            canonicalAction, names);
                    }
                    catch (ArgumentException ex)
                    {
                        throw new ArgumentException($"{ex.Message} Supplied MCP parameters: {string.Join(", ", names)}.");
                    }
                    if (tool.ProtocolTool.Name == "worksheet" && canonicalAction is "copy-to-file" or "move-to-file"
                        && arguments.ContainsKey("workbook_session_id"))
                        throw new ArgumentException("workbook_session_id is not used by atomic cross-file worksheet actions.");
                }
            }
            catch (ArgumentException ex)
            {
                return ValueTask.FromResult(ExcelToolsBase.CreateToolResult(
                    ExcelToolsBase.SerializeToolError(tool.ProtocolTool.Name, null, ex), isError: true));
            }

            return next(request, cancellationToken);
        };

    private static void ValidateValueKind(string name, JsonElement value, JsonElement schema)
    {
        if (!schema.TryGetProperty("type", out var type))
            return;

        var types = type.ValueKind == JsonValueKind.Array
            ? type.EnumerateArray().Select(t => t.GetString()).ToArray()
            : [type.GetString()];
        var valid = types.Any(t => t switch
        {
            "null" => value.ValueKind == JsonValueKind.Null,
            "string" => value.ValueKind == JsonValueKind.String,
            "boolean" => value.ValueKind is JsonValueKind.True or JsonValueKind.False,
            "integer" => value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out _),
            "number" => value.ValueKind == JsonValueKind.Number,
            "array" => value.ValueKind == JsonValueKind.Array,
            "object" => value.ValueKind == JsonValueKind.Object,
            _ => true
        });
        if (!valid)
            throw new ArgumentException($"Parameter '{name}' must have type {string.Join(" or ", types)}.");
    }
}
