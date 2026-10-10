using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Utilities;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

[McpServerToolType]
public static partial class ExcelFileTool
{
    /// <summary>
    /// Open/create workbooks and manage their sessions.
    /// Reuse the intended workbook's existing session or open/create one, then operate and close when authorized.
    /// Open/create return workbook_session_id; pass it to session-based tools. Create requires an existing directory.
    /// Close defaults to save:false (discard edits); set save:true to save. Wait for canClose before closing,
    /// and confirm before closing a visible window unless already authorized.
    /// Normal server shutdown attempts to save open sessions; crashes and forced cleanup may lose edits.
    /// Protected files require show:true for authentication; Excel determines editing rights.
    /// Open accepts direct SharePoint/OneDrive for Business HTTPS workbook URLs, including ?web=1.
    /// URL opening requires show:true; browser pages, sharing links and arbitrary web URLs are not supported.
    /// Remote AutoSave is disabled so explicit save/close semantics apply.
    /// Inspect workbook_read get-info readOnly before editing. Workbook changes reject read-only access.
    /// Failed saves retain the open session and unsaved changes.
    /// canClose is false while Excel has a modal dialog open, is busy, refreshing, or its state cannot be confirmed,
    /// even with activeOperations:0. Close with either save value is blocked until Excel is ready.
    /// Open/create default to 120 seconds. Cancellation is not undo; inspect the session state before continuing.
    /// </summary>
    /// <param name="action">The file operation to perform. close with save:false discards all unsaved edits, including earlier work; there is no tool-level undo.</param>
    /// <param name="file_path">Full Windows workbook path, or a direct SharePoint/OneDrive for Business HTTPS workbook URL for open. URLs support .xlsx/.xlsm/.xlsb/.xls and optional ?web=1 or ?web=0. Create requires a Windows path and supports .xlsx/.xlsm. Use a supplied location or discover the matching session; ask if the intended file is unclear.</param>
    /// <param name="workbook_session_id">Session ID returned by open/create or listed by this server. Required for close.</param>
    /// <param name="save">Save before close; otherwise discard unsaved changes. Only valid for close.</param>
    /// <param name="show">Show Excel. Only valid for open/create; protected files may force visible authentication.</param>
    /// <param name="timeout_seconds">Timeout for open/create, in seconds (10-3600). Also sets the session operation timeout.</param>
    [McpServerTool(Name = "file", Title = "File Operations", Destructive = true,
        UseStructuredContent = true, OutputSchemaType = typeof(FileToolOutputSchema))]
    [McpMeta("category", "session")]
    [McpMeta("requiresSession", false)]
    public static partial Task<CallToolResult> ExcelFile(
        FileWriteAction action,
        ServiceBridge.ServiceBridge bridge,
        [McpActionParameter("open", Required = true), McpActionParameter("create", Required = true)]
        [DefaultValue(null)] string? file_path,
        [McpActionParameter("close", Required = true)]
        [DefaultValue(null)] string? workbook_session_id,
        [McpActionParameter("close"), DefaultValue(false)] bool save,
        [McpActionParameter("open"), McpActionParameter("create"), DefaultValue(false)] bool show,
        [McpActionParameter("open"), McpActionParameter("create"), DefaultValue(120)] int timeout_seconds,
        CancellationToken cancellationToken = default) =>
        ExecuteFileToolActionAsync(
            "file",
            action switch
            {
                FileWriteAction.Open => FileAction.Open,
                FileWriteAction.Create => FileAction.Create,
                FileWriteAction.Close => FileAction.Close,
                _ => throw new ArgumentOutOfRangeException(nameof(action))
            },
            bridge, file_path, workbook_session_id, save, show, timeout_seconds, cancellationToken);

    /// <summary>List workbook sessions with live canClose, excelState and blockingReason, and validate a workbook path without opening an editable session. dialogOpen means an Excel-owned modal window is visible: ask the user to check Excel for a prompt, which may require sign-in. It does not identify the dialog type or prove a query has stopped. activeOperations:0 does not prove a background refresh has finished.</summary>
    /// <remarks>
    /// Test defaults to 120 seconds and validates ordinary files through a temporary read-only Excel open.
    /// IRM/AIP files require visible authentication; Excel determines editing rights, not protection detection.
    /// Inspect canOpen, isIrmProtected, willOpenReadOnly, and requiresVisibleSession.
    /// willOpenReadOnly:false does not guarantee editing rights; inspect workbook_read get-info readOnly after opening.
    /// Test does not bypass authentication.
    /// SharePoint URLs return an interactive-validation requirement without opening Excel.
    /// Remote existence, size and IRM protection cannot be established by local preflight;
    /// false exists/isIrmProtected values do not prove absence or lack of protection.
    /// </remarks>
    /// <param name="action">List sessions or test whether a workbook can be opened.</param>
    /// <param name="file_path">Full Windows workbook path or direct SharePoint/OneDrive for Business HTTPS workbook URL. Required for test.</param>
    /// <param name="timeout_seconds">Timeout for test, in seconds (10-3600).</param>
    [McpServerTool(Name = "file_read", Title = "Read-Only File Operations", ReadOnly = true,
        Destructive = false, UseStructuredContent = true, OutputSchemaType = typeof(FileToolOutputSchema))]
    [McpMeta("category", "session")]
    [McpMeta("requiresSession", false)]
    public static partial Task<CallToolResult> ExcelFileRead(
        FileReadAction action,
        ServiceBridge.ServiceBridge bridge,
        [McpActionParameter("test", Required = true), DefaultValue(null)] string? file_path,
        [McpActionParameter("test"), DefaultValue(120)] int timeout_seconds = 120,
        CancellationToken cancellationToken = default) =>
        ExecuteFileToolActionAsync(
            "file_read",
            action switch
            {
                FileReadAction.List => FileAction.List,
                FileReadAction.Test => FileAction.Test,
                _ => throw new ArgumentOutOfRangeException(nameof(action))
            },
            bridge, file_path, null, save: false, show: false,
            timeout_seconds: timeout_seconds, cancellationToken: cancellationToken);

    private static Task<CallToolResult> ExecuteFileToolActionAsync(
        string toolName,
        FileAction action,
        ServiceBridge.ServiceBridge bridge,
        string? file_path,
        string? workbook_session_id,
        bool save,
        bool show,
        int timeout_seconds,
        CancellationToken cancellationToken) =>
        ExcelToolsBase.ExecuteToolActionAsync(toolName, action.ToActionString(), async () =>
        {
            if (timeout_seconds is < 10 or > 3600)
                throw new ArgumentException("timeout_seconds must be between 10 and 3600 seconds.");

            if (action is FileAction.Open or FileAction.Create or FileAction.Test && string.IsNullOrWhiteSpace(file_path))
                throw new ArgumentException($"file_path is required for '{action.ToActionString()}' action.");
            if (action == FileAction.Close && string.IsNullOrWhiteSpace(workbook_session_id))
                throw new ArgumentException("workbook_session_id is required for file 'close'.");

            var isRemoteOpen = action == FileAction.Open
                && FilePathValidation.IsRemoteWorkbook(file_path!);
            if (action is FileAction.Open or FileAction.Create && !isRemoteOpen)
            {
                var pathError = ExcelToolsBase.ValidateWindowsPath(file_path);
                if (pathError is not null)
                    return pathError;
            }
            if (action == FileAction.Open && !isRemoteOpen && !File.Exists(file_path))
            {
                return JsonSerializer.Serialize(new
                {
                    success = false,
                    errorMessage = $"File not found: {file_path}",
                    errorCategory = "NotFound",
                    filePath = file_path,
                    isError = true
                }, ExcelToolsBase.JsonOptions);
            }

            var response = action switch
            {
                FileAction.List => await bridge.SendAsync("session.list", cancellationToken: cancellationToken),
                FileAction.Close => await bridge.SendAsync("session.close", workbook_session_id, new { save }, cancellationToken: cancellationToken),
                FileAction.Open or FileAction.Create => await bridge.SendAsync(
                    $"session.{action.ToActionString()}", args: new { filePath = file_path, show, timeoutSeconds = timeout_seconds },
                    timeoutSeconds: timeout_seconds, cancellationToken: cancellationToken),
                FileAction.Test => await bridge.SendAsync("session.test", args: new { filePath = file_path, timeoutSeconds = timeout_seconds },
                    timeoutSeconds: timeout_seconds, cancellationToken: cancellationToken),
                _ => throw new ArgumentException($"Unknown file action: {action}.")
            };

            if (!response.Success)
                return ExcelToolsBase.SerializeServiceResponse(response, file_path);

            // session.close is a void Service command; its successful acknowledgement is authoritative.
            if (action == FileAction.Close)
                return JsonSerializer.Serialize(new { success = true, workbook_session_id, saved = save }, ExcelToolsBase.JsonOptions);

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
                    session["workbook_session_id"] = sessionId;
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
                workbook_session_id = id,
                filePath = document.RootElement.GetProperty("filePath").GetString()
            }, ExcelToolsBase.JsonOptions);
        }, cancellationToken);
}
