using System.ComponentModel;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Attributes;

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
    /// Protected files may need visible authentication; Excel determines editing rights.
    /// Open passes Windows paths and URLs to Excel, which determines supported formats and access.
    /// URL parameters are passed unchanged. Use show:true when you need to answer Excel sign-in or permission prompts.
    /// Open and attach preserve and report AutoSave. If enabled, Excel may save edits automatically; close without saving cannot undo already-saved edits.
    /// Explicitly change AutoSave with workbook set-auto-save enabled. Enabling can immediately save pending edits.
    /// Inspect workbook_read get-info readOnly before editing. Workbook changes reject read-only access.
    /// Failed saves retain the open session and unsaved changes.
    /// Close returns closeTimingsMs with measured step durations. Some steps include others; do not add them together.
    /// Save persists the current workbook without closing it. Timeouts preserve the owned session,
    /// request that its Excel window be shown, and report the result to the caller.
    /// canClose is false while Excel has a modal dialog open, is busy, refreshing, or its state cannot be confirmed,
    /// even with activeOperations:0. Confirmed blocking dialogs reveal Excel and report generic dialogOpen with operationPending true.
    /// Close with either save value is blocked until Excel is ready.
    /// Open/create default to 120 seconds. Cancellation is not undo; inspect the session state before continuing.
    /// Opening disables VBA macros. enable-macros requires approved:true and allow_reopen:true, a saved local workbook opened by this session, and no neighboring workbooks. Opening event code can run on reopen.
    /// enable-editing and enable-connections require approved:true and allow_remembered_trust:true after the user approves that Excel may remember trust for this file. Editing does not also enable VBA macros or update external links. Connection enabling affects all workbook connections and can contact data sources or start refresh.
    /// Company restrictions remain in force. Approval arguments are explicit opt-ins, not a built-in user confirmation dialog.
    /// </summary>
    /// <param name="action">The file operation to perform. close with save:false discards all unsaved edits, including earlier work; there is no tool-level undo.</param>
    /// <param name="file_path">Full Windows path or URL for open. Excel determines what it can open. Create requires a Windows path and supports .xlsx/.xlsm. Use a supplied location or discover the matching session; ask if the intended file is unclear.</param>
    /// <param name="workbook_session_id">Session ID returned by open/create/attach or listed by this server. Required for close, save, enable-editing, enable-connections, and enable-macros.</param>
    /// <param name="save">Save before close; otherwise discard unsaved changes. Only valid for close.</param>
    /// <param name="show">Show Excel. Only valid for open/create; protected files may force visible authentication.</param>
    /// <param name="timeout_seconds">Timeout for open/create, in seconds (10-3600). Also sets the session operation timeout.</param>
    /// <param name="open_workbook_id">Identifier returned by file_read discover. Required for attach.</param>
    /// <param name="approved">Explicit approval of the selected workbook security action. Only valid for enable-editing, enable-connections, enable-macros; default false.</param>
    /// <param name="allow_remembered_trust">Approval that Excel can remember this file as trusted. Required true for enable-editing and enable-connections; default false.</param>
    /// <param name="allow_reopen">Separate approval to close and reopen the saved workbook, allowing opening event macros to run. Required true for enable-macros; default false. Never saves or discards unsaved work.</param>
    [McpServerTool(Name = "file", Title = "File Operations", Destructive = true)]
    [McpMeta("category", "session")]
    [McpMeta("requiresSession", false)]
    public static partial Task<CallToolResult> ExcelFile(
        FileWriteAction action,
        ServiceBridge.ServiceBridge bridge,
        [McpActionParameter("open", Required = true), McpActionParameter("create", Required = true)]
        [DefaultValue(null)] string? file_path,
        [McpActionParameter("close", Required = true), McpActionParameter("save", Required = true)]
        [McpActionParameter("enable-editing", Required = true), McpActionParameter("enable-connections", Required = true), McpActionParameter("enable-macros", Required = true)]
        [DefaultValue(null)] string? workbook_session_id,
        [McpActionParameter("close"), DefaultValue(false)] bool save,
        [McpActionParameter("open"), McpActionParameter("create"), DefaultValue(false)] bool show,
        [McpActionParameter("open"), McpActionParameter("create"), DefaultValue(120)] int timeout_seconds,
        [McpActionParameter("attach", Required = true), DefaultValue(null)] string? open_workbook_id = null,
        [McpActionParameter("enable-editing"), McpActionParameter("enable-connections"), McpActionParameter("enable-macros"), DefaultValue(false)] bool approved = false,
        [McpActionParameter("enable-editing"), McpActionParameter("enable-connections"), DefaultValue(false)] bool allow_remembered_trust = false,
        [McpActionParameter("enable-macros"), DefaultValue(false)] bool allow_reopen = false,
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
                FileWriteAction.EnableEditing => FileAction.EnableEditing,
                FileWriteAction.EnableConnections => FileAction.EnableConnections,
                FileWriteAction.EnableMacros => FileAction.EnableMacros,
                _ => throw new ArgumentOutOfRangeException(nameof(action))
            },
            bridge, file_path, workbook_session_id, save, show, timeout_seconds, cancellationToken, open_workbook_id,
            approved, allow_remembered_trust, allow_reopen);

    /// <summary>Discover already-open workbooks in the current user's Windows session without reopening them, list managed sessions, test a workbook path, or inspect a selected session's opening restrictions with notices. Notices never enable content and return complete:false because not every notification bar is readable. Discover returns workbookId values for file attach, sessionId for workbooks already connected to this host, and complete:false with errors when any application or workbook could not be inspected. Discovery can attach Protected View workbooks when a normal controlling Excel window is available; restricted-only windows report an error. List includes canClose, excelState, operationPending and blockers. Confirmed blocking dialogs reveal Excel and report readable dialog text in blockingMessage. Ask the user to respond in Excel; input values are not read. activeOperations:0 does not prove refresh has finished.</summary>
    /// <remarks>
    /// Test defaults to 120 seconds and validates ordinary files through a temporary read-only Excel open.
    /// IRM/AIP files require visible authentication; Excel determines editing rights, not protection detection.
    /// Inspect canOpen, isIrmProtected, willOpenReadOnly, and requiresVisibleSession.
    /// willOpenReadOnly:false does not guarantee editing rights; inspect workbook_read get-info readOnly after opening.
    /// Test does not bypass authentication.
    /// URLs report unavailable local file details without opening Excel; use open to check access.
    /// Remote existence, size and IRM protection cannot be established by local preflight;
    /// false exists/isIrmProtected values do not prove absence or lack of protection.
    /// </remarks>
    /// <param name="action">Discover already-open workbooks, list sessions, test whether a workbook can be opened, or inspect a selected session's opening notices.</param>
    /// <param name="file_path">Full Windows path or URL. Required for test.</param>
    /// <param name="timeout_seconds">Timeout for test, in seconds (10-3600).</param>
    /// <param name="workbook_session_id">Session to inspect for notices. Attach a discovered Protected View workbook first. Reports verified restrictions and explicit inspection gaps; complete:false does not mean content is enabled.</param>
    [McpServerTool(Name = "file_read", Title = "Read-Only File Operations", ReadOnly = true,
        Destructive = false)]
    [McpMeta("category", "session")]
    [McpMeta("requiresSession", false)]
    public static partial Task<CallToolResult> ExcelFileRead(
        FileReadAction action,
        ServiceBridge.ServiceBridge bridge,
        [McpActionParameter("test", Required = true), DefaultValue(null)] string? file_path,
        [McpActionParameter("test"), DefaultValue(120)] int timeout_seconds = 120,
        [McpActionParameter("notices", Required = true), DefaultValue(null)] string? workbook_session_id = null,
        CancellationToken cancellationToken = default) =>
        ExecuteFileReadAsync(action, bridge, file_path, timeout_seconds, workbook_session_id, cancellationToken);

    private static Task<CallToolResult> ExecuteFileReadAsync(
        FileReadAction action, ServiceBridge.ServiceBridge bridge, string? filePath,
        int timeoutSeconds, string? sessionId, CancellationToken cancellationToken)
    {
        var fileAction = action switch
        {
            FileReadAction.List => FileAction.List,
            FileReadAction.Test => FileAction.Test,
            FileReadAction.Discover => FileAction.Discover,
            FileReadAction.Notices => FileAction.Notices,
            _ => throw new ArgumentOutOfRangeException(nameof(action))
        };
        return ExecuteFileToolActionAsync(
            "file_read",
            fileAction, bridge, filePath, sessionId, save: false, show: false,
            timeout_seconds: timeoutSeconds, cancellationToken: cancellationToken);
    }

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
        string? open_workbook_id = null,
        bool approved = false,
        bool allow_remembered_trust = false,
        bool allow_reopen = false) =>
        ExcelToolsBase.ExecuteToolActionAsync(toolName, action.ToActionString(), async () =>
        {
            if (timeout_seconds is < 10 or > 3600)
                throw new ArgumentException("timeout_seconds must be between 10 and 3600 seconds.");

            if (action is FileAction.Open or FileAction.Create or FileAction.Test && string.IsNullOrWhiteSpace(file_path))
                throw new ArgumentException($"file_path is required for '{action.ToActionString()}' action.");
            if (action is FileAction.Close or FileAction.Save or FileAction.Notices
                or FileAction.EnableEditing or FileAction.EnableConnections or FileAction.EnableMacros
                && string.IsNullOrWhiteSpace(workbook_session_id))
                throw new ArgumentException($"workbook_session_id is required for file '{action.ToActionString()}'.");
            if (action == FileAction.Attach && string.IsNullOrWhiteSpace(open_workbook_id))
                throw new ArgumentException("open_workbook_id from file_read discover is required for attach.");

            if (action == FileAction.Create)
            {
                var pathError = ExcelToolsBase.ValidateWindowsPath(file_path);
                if (pathError is not null)
                    return pathError;
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
                FileAction.Notices => await bridge.SendAsync("session.notices", workbook_session_id, cancellationToken: cancellationToken),
                FileAction.EnableEditing or FileAction.EnableConnections => await bridge.SendAsync(
                    $"session.{action.ToActionString()}", workbook_session_id,
                    new { approved, allowRememberedTrust = allow_remembered_trust }, cancellationToken: cancellationToken),
                FileAction.EnableMacros => await bridge.SendAsync("session.enable-macros", workbook_session_id,
                    new { approved, allowReopen = allow_reopen }, cancellationToken: cancellationToken),
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

            var command = ServiceRegistry.GetMcpCommand(toolName, action.ToActionString());
            var outcome = ExcelToolsBase.ReadServiceResponse(response, command);
            if (!outcome.Success)
            {
                if (file_path is not null)
                    outcome.Error!["filePath"] = file_path;
                return outcome.ToToolJson();
            }

            var result = System.Text.Json.Nodes.JsonNode.Parse(outcome.ToToolJson())!.AsObject();
            var id = result["sessionId"]?.GetValue<string>();
            if (action is FileAction.Open or FileAction.Create or FileAction.Attach && string.IsNullOrWhiteSpace(id))
                throw new InvalidOperationException("Workbook startup returned no session ID.");
            if (id is not null)
            {
                result.Remove("sessionId");
                result["workbook_session_id"] = id;
            }
            return result.ToJsonString(ExcelToolsBase.JsonOptions);
        }, cancellationToken);
}
