using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

[McpServerToolType]
public static partial class ExcelFileTool
{
    internal static void ValidateActionParameters(string action, IEnumerable<string> names)
    {
        string[] allowed = action switch
        {
            "open" or "create" => ["action", "path", "show", "timeout_seconds"],
            "test" => ["action", "path", "timeout_seconds"],
            "close" => ["action", "session_id", "save"],
            "list" => ["action"],
            _ => throw new ArgumentException("Unknown file action.")
        };
        var invalid = names.Except(allowed, StringComparer.Ordinal).ToArray();
        if (invalid.Length > 0)
            throw new ArgumentException($"Parameter(s) {string.Join(", ", invalid)} are not valid for file '{action}'.");
    }

    /// <summary>
    /// Open/create workbooks and manage their sessions.
    /// Workflow: list and match the intended workbook -> reuse its session or open/create -> operate ->
    /// list and check that session's canClose -> close when authorized with explicit save:true or save:false.
    /// Open/create and list entries return session_id; pass it to session-based tools. Create requires an existing directory.
    /// Close defaults to save:false (discard edits); set save:true to save. Wait for canClose before closing,
    /// and confirm before closing a visible window unless already authorized.
    /// Normal server shutdown attempts to save open sessions; crashes and forced cleanup may lose edits.
    /// Open/create/test default to 120 seconds. Cancellation is not undo; inspect list before continuing.
    /// Test checks path, extension, access, and IRM/AIP signals without opening or inspecting workbook contents.
    /// Inspect preflightPassed, isIrmProtected, willOpenReadOnly, and requiresVisibleSession;
    /// isValid and canOpen remain false until Excel opens the workbook. Test does not bypass authentication.
    /// </summary>
    /// <param name="action">The file operation to perform. close with save:false discards all unsaved edits, including earlier work; there is no tool-level undo.</param>
    /// <param name="path">Absolute native workbook path. Required for open, create, test. Create supports .xlsx and, on Windows, .xlsm. Use a supplied path or discover the matching session; ask if the intended file is unclear.</param>
    /// <param name="session_id">Session ID returned by open/create or listed by this server. Required for close.</param>
    /// <param name="save">Save before close; otherwise discard unsaved changes. Only valid for close.</param>
    /// <param name="show">Show Excel. Only valid for open/create; protected files may force visible authentication.</param>
    /// <param name="timeout_seconds">Timeout for open/create/test, in seconds (10-3600). Open/create also sets the session operation timeout.</param>
    [McpServerTool(Name = "file", Title = "File Operations", Destructive = true,
        UseStructuredContent = true, OutputSchemaType = typeof(FileToolOutputSchema))]
    [McpMeta("category", "session")]
    [McpMeta("requiresSession", false)]
    public static partial Task<CallToolResult> ExcelFile(
        FileAction action,
        ServiceBridge.ServiceBridge bridge,
        [DefaultValue(null)] string? path,
        [DefaultValue(null)] string? session_id,
        [DefaultValue(false)] bool save,
        [DefaultValue(false)] bool show,
        [DefaultValue(120)] int timeout_seconds,
        CancellationToken cancellationToken = default) =>
        ExcelToolsBase.ExecuteToolActionAsync("file", action.ToActionString(), async () =>
        {
            if (timeout_seconds is < 10 or > 3600)
                throw new ArgumentException("timeout_seconds must be between 10 and 3600 seconds.");

            if (action is FileAction.Open or FileAction.Create or FileAction.Test && string.IsNullOrWhiteSpace(path))
                throw new ArgumentException($"path is required for '{action.ToActionString()}' action.");
            if (action == FileAction.Close && string.IsNullOrWhiteSpace(session_id))
                throw new ArgumentException(SessionIdentityFilter.ErrorMessage);

            if (action is FileAction.Open or FileAction.Create)
            {
                var pathError = ExcelToolsBase.ValidateAbsolutePath(path);
                if (pathError is not null)
                    return pathError;
            }
            if (action == FileAction.Open && !File.Exists(path))
            {
                return JsonSerializer.Serialize(new
                {
                    success = false,
                    errorMessage = $"File not found: {path}",
                    errorCategory = "NotFound",
                    filePath = path,
                    isError = true
                }, ExcelToolsBase.JsonOptions);
            }

            var response = action switch
            {
                FileAction.List => await bridge.SendAsync("session.list", cancellationToken: cancellationToken),
                FileAction.Close => await bridge.SendAsync("session.close", session_id, new { save }, cancellationToken: cancellationToken),
                FileAction.Open or FileAction.Create => await bridge.SendAsync(
                    $"session.{action.ToActionString()}", args: new { filePath = path, show, timeoutSeconds = timeout_seconds },
                    timeoutSeconds: timeout_seconds, cancellationToken: cancellationToken),
                FileAction.Test => await bridge.SendAsync("session.test", args: new { filePath = path, timeoutSeconds = timeout_seconds },
                    timeoutSeconds: timeout_seconds, cancellationToken: cancellationToken),
                _ => throw new ArgumentException($"Unknown file action: {action}.")
            };

            if (!response.Success)
                return ExcelToolsBase.SerializeServiceResponse(response, path);

            // session.close is a void Service command; its successful acknowledgement is authoritative.
            if (action == FileAction.Close)
                return JsonSerializer.Serialize(new { success = true, session_id, saved = save }, ExcelToolsBase.JsonOptions);

            var result = response.Result
                ?? throw new InvalidOperationException("File operation returned no result.");
            if (action == FileAction.List)
            {
                var list = JsonNode.Parse(result)?.AsObject()
                    ?? throw new InvalidOperationException("Session listing returned no object.");
                foreach (var node in list["sessions"]?.AsArray()
                    ?? throw new InvalidOperationException("Session listing returned no sessions array."))
                {
                    var session = node?.AsObject()
                        ?? throw new InvalidOperationException("Session listing returned an invalid entry.");
                    var sessionId = session["sessionId"]?.GetValue<string>();
                    if (string.IsNullOrWhiteSpace(sessionId))
                        throw new InvalidOperationException("Session listing returned no session ID.");
                    session.Remove("sessionId");
                    session["session_id"] = sessionId;
                }
                return list.ToJsonString(ExcelToolsBase.JsonOptions);
            }
            if (action is not (FileAction.Open or FileAction.Create))
                return result;

            using var document = JsonDocument.Parse(result);
            var id = document.RootElement.GetProperty("sessionId").GetString();
            if (string.IsNullOrWhiteSpace(id))
                throw new InvalidOperationException("Workbook startup returned no session ID.");
            return JsonSerializer.Serialize(new
            {
                success = true,
                session_id = id,
                filePath = document.RootElement.GetProperty("filePath").GetString()
            }, ExcelToolsBase.JsonOptions);
        }, cancellationToken);
}
