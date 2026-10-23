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
    /// Normal shutdown closes saved owned workbooks and leaves unsaved ones open and visible without saving or discarding. Attached workbooks stay untouched.
    /// Protected files require show:true for authentication; Excel determines editing rights.
    /// Open accepts direct SharePoint/OneDrive for Business HTTPS workbook URLs, including ?web=1.
    /// URL opening requires show:true; browser pages, sharing links and arbitrary web URLs are not supported.
    /// Open and attach preserve and report AutoSave. If enabled, Excel may save edits automatically; close without saving cannot undo already-saved edits.
    /// Explicitly change AutoSave with workbook set-auto-save enabled. Enabling can immediately save pending edits.
    /// Inspect workbook_read get-info readOnly before editing. Workbook changes reject read-only access.
    /// Failed saves retain the open session and unsaved changes.
    /// Save persists the current workbook without closing it. Timeouts preserve the owned session,
    /// request that its Excel window be shown, and report the result to the caller.
    /// canClose is false while Excel has a modal dialog open, is busy, refreshing, or its state cannot be confirmed,
    /// even with activeOperations:0. Confirmed blocking dialogs reveal Excel and report generic dialogOpen with operationPending true.
    /// Close with either save value is blocked until Excel is ready.
    /// Open/create default to 120 seconds. Cancellation is not undo; inspect the session state before continuing.
    /// </summary>
    /// <param name="action">The file operation to perform. close with save:false discards all unsaved edits, including earlier work; there is no tool-level undo.</param>
    /// <param name="file_path">Full Windows workbook path, or a direct SharePoint/OneDrive for Business HTTPS workbook URL for open. URLs support .xlsx/.xlsm/.xlsb/.xls and optional ?web=1 or ?web=0. Create requires a Windows path and supports .xlsx/.xlsm. Use a supplied location or discover the matching session; ask if the intended file is unclear.</param>
    /// <param name="workbook_session_id">Session ID returned by open/create or listed by this server. Required for close/save.</param>
    /// <param name="save">Save before close; otherwise discard unsaved changes. Only valid for close.</param>
    /// <param name="show">Show Excel. Only valid for open/create; protected files may force visible authentication.</param>
    /// <param name="timeout_seconds">Timeout for open/create, in seconds (10-3600). Also sets the session operation timeout.</param>
    /// <param name="open_workbook_id">Identifier returned by file_read discover. Required for attach.</param>
    [McpServerTool(Name = "file", Title = "File Operations", Destructive = true,
        UseStructuredContent = true, OutputSchemaType = typeof(FileToolOutputSchema))]
    [McpMeta("category", "session")]
    [McpMeta("requiresSession", false)]
    public static partial Task<CallToolResult> ExcelFile(
        FileWriteAction action,
        ServiceBridge.ServiceBridge bridge,
        [McpActionParameter("open", Required = true), McpActionParameter("create", Required = true)]
        [DefaultValue(null)] string? file_path,
        [McpActionParameter("close", Required = true), McpActionParameter("save", Required = true)]
        [DefaultValue(null)] string? workbook_session_id,
        [McpActionParameter("close"), DefaultValue(false)] bool save,
        [McpActionParameter("open"), McpActionParameter("create"), DefaultValue(false)] bool show,
        [McpActionParameter("open"), McpActionParameter("create"), DefaultValue(120)] int timeout_seconds,
        [McpActionParameter("attach", Required = true), DefaultValue(null)] string? open_workbook_id = null,
        CancellationToken cancellationToken = default) =>
        ExecuteFileToolActionAsync(
            "file",
            action switch
            {
                FileWriteAction.Open => FileAction.Open,
                FileWriteAction.Create => FileAction.Create,
                FileWriteAction.Close => FileAction.Close,
                FileWriteAction.Save => FileAction.Save,
                FileWriteAction.Attach => FileAction.Attach,
                _ => throw new ArgumentOutOfRangeException(nameof(action))
            },
            bridge, file_path, workbook_session_id, save, show, timeout_seconds, cancellationToken, open_workbook_id);

    /// <summary>Discover already-open workbooks in the current user's Windows session without reopening them, list managed sessions, or test a workbook path. Discover returns workbookId values for file attach, sessionId for workbooks already connected to this host, and complete:false with errors when any application or workbook could not be inspected. List includes canClose, excelState, operationPending and blockers. Confirmed blocking dialogs reveal Excel and report readable dialog text in blockingMessage. Ask the user to respond in Excel; input values are not read. activeOperations:0 does not prove refresh has finished.</summary>
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
    /// <param name="action">Discover already-open workbooks, list sessions, or test whether a workbook can be opened.</param>
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
                FileReadAction.Discover => FileAction.Discover,
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
        CancellationToken cancellationToken,
        string? open_workbook_id = null) =>
        ExcelToolsBase.ExecuteToolActionAsync(toolName, action.ToActionString(), async () =>
        {
            if (timeout_seconds is < 10 or > 3600)
                throw new ArgumentException("timeout_seconds must be between 10 and 3600 seconds.");

            if (action is FileAction.Open or FileAction.Create or FileAction.Test && string.IsNullOrWhiteSpace(file_path))
                throw new ArgumentException($"file_path is required for '{action.ToActionString()}' action.");
            if (action is FileAction.Close or FileAction.Save && string.IsNullOrWhiteSpace(workbook_session_id))
                throw new ArgumentException($"workbook_session_id is required for file '{action.ToActionString()}'.");
            if (action == FileAction.Attach && string.IsNullOrWhiteSpace(open_workbook_id))
                throw new ArgumentException("open_workbook_id from file_read discover is required for attach.");

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
                FileAction.Discover => await bridge.SendAsync("session.discover", cancellationToken: cancellationToken),
                FileAction.Attach => await bridge.SendAsync(
                    "session.attach", args: new { workbookId = open_workbook_id },
                    cancellationToken: cancellationToken),
                FileAction.Close => await bridge.SendAsync("session.close", workbook_session_id, new { save }, cancellationToken: cancellationToken),
                FileAction.Save => await bridge.SendAsync("session.save", workbook_session_id, cancellationToken: cancellationToken),
                FileAction.Open or FileAction.Create => await bridge.SendAsync(
                    $"session.{action.ToActionString()}", args: new
                    {
                        filePath = file_path,
                        show,
                        timeoutSeconds = timeout_seconds
                    },
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
            if (action == FileAction.Save)
                return JsonSerializer.Serialize(new { success = true, workbook_session_id, saved = true }, ExcelToolsBase.JsonOptions);

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
            if (action == FileAction.Discover)
                return result;
            if (action == FileAction.Attach)
            {
                using var attachDocument = JsonDocument.Parse(result);
                var attach = attachDocument.RootElement;
                return JsonSerializer.Serialize(new
                {
                    success = true,
                    workbook_session_id = attach.GetProperty("sessionId").GetString(),
                    filePath = attach.TryGetProperty("filePath", out var filePath)
                        && filePath.ValueKind != JsonValueKind.Null
                        ? filePath.GetString()
                        : null,
                    name = attach.GetProperty("name").GetString(),
                    autoSaveOn = attach.TryGetProperty("autoSaveOn", out var autoSaveOn)
                        && autoSaveOn.ValueKind != JsonValueKind.Null
                        ? autoSaveOn.GetBoolean()
                        : (bool?)null,
                    warning = attach.GetProperty("warning").GetString()
                }, ExcelToolsBase.JsonOptions);
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
                filePath = document.RootElement.GetProperty("filePath").GetString(),
                autoSaveOn = document.RootElement.TryGetProperty("autoSaveOn", out var startupAutoSave)
                    && startupAutoSave.ValueKind != JsonValueKind.Null ? startupAutoSave.GetBoolean() : (bool?)null,
                warning = document.RootElement.TryGetProperty("warning", out var startupWarning)
                    ? startupWarning.GetString() : null
            }, ExcelToolsBase.JsonOptions);
        }, cancellationToken);
}
