using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel worksheet management tool for MCP server.
/// Handles both session-based operations (list, create, rename, delete, move, copy)
/// and atomic cross-file operations (copy-to-file, move-to-file).
/// </summary>
[McpServerToolType]
public static partial class ExcelWorksheetTool
{
    /// <summary>
    /// Worksheet lifecycle: create, rename, copy, delete, move.
    /// RENAME: Use old_name + new_name.
    /// ATOMIC OPERATIONS: copy-to-file and move-to-file don't require a session (open/close automatically).
    /// POSITIONING: Use before OR after (not both) to place sheet relative to another.
    /// Use worksheet_style for tab colors and visibility.
    /// </summary>
    /// <param name="action">The action to perform</param>
    /// <param name="session_id">Session ID from file 'open' action (required for: list, create, rename, delete, move, copy. Not required for: copy-to-file, move-to-file)</param>
    /// <param name="sheet_name">Name of the worksheet (required for: create, delete, move)</param>
    /// <param name="old_name">Current name of the worksheet (required for: rename)</param>
    /// <param name="source_name">Name of the source worksheet (required for: copy)</param>
    /// <param name="target_name">Name for the copied worksheet (required for: copy)</param>
    /// <param name="new_name">New name for the worksheet (required for: rename)</param>
    /// <param name="file_path">Optional file path when batch contains multiple workbooks</param>
    /// <param name="source_file">Full path to the source workbook (required for: copy-to-file, move-to-file)</param>
    /// <param name="source_sheet">Name of the sheet to copy (required for: copy-to-file, move-to-file)</param>
    /// <param name="target_file">Full path to the target workbook (required for: copy-to-file, move-to-file)</param>
    /// <param name="target_sheet_name">Optional: New name for the copied sheet (default: keeps original name)</param>
    /// <param name="before_sheet">Optional: Position before this sheet</param>
    /// <param name="after_sheet">Optional: Position after this sheet</param>
    /// <param name="review_only">Return an exact mutation plan without changing Excel.</param>
    /// <param name="review_id">Bound review identifier returned by a prior review-only request.</param>
    /// <param name="checkpoint">Create a recoverable workbook checkpoint immediately before a mutation.</param>
    /// <param name="idempotency_key">Client-generated key for safely retrying a same-workbook mutation.</param>
    /// <param name="cancellationToken">Cancellation token for the tool request.</param>
    [McpServerTool(Name = "worksheet", Title = "Worksheet Operations", Destructive = true)]
    [McpMeta("category", "structure")]
    [McpMeta("requiresSession", false)]  // Session is optional - depends on the action
    [Description("Worksheet lifecycle: create, rename, copy, delete, move. RENAME: Use old_name plus new_name. ATOMIC OPERATIONS: copy-to-file and move-to-file don't require a session (open/close automatically). POSITIONING: Use before OR after (not both) to place sheet relative to another. Use worksheet_style for tab colors and visibility.")]
    public static string ExcelWorksheet(
        [Description("The action to perform")] SheetAction action,
        [Description(
            "Session ID from file 'open' or 'create'. Required for same-workbook actions: list, create, rename, delete, move, and copy. Not used by copy-to-file or move-to-file.")]
        string? session_id = null,
        [Description(
            "Worksheet name for create, delete, and move.")]
        string? sheet_name = null,
        [Description(
            "Current worksheet name for rename.")]
        string? old_name = null,
        [Description(
            "Source worksheet name for copy within the same workbook.")]
        string? source_name = null,
        [Description(
            "Target worksheet name for copy within the same workbook.")]
        string? target_name = null,
        [Description(
            "New worksheet name for rename.")]
        string? new_name = null,
        [Description(
            "Optional workbook path when the current batch session has multiple open workbooks.")]
        string? file_path = null,
        [Description("Source workbook path for copy-to-file and move-to-file.")]
        string? source_file = null,
        [Description(
            "Source worksheet name for copy-to-file and move-to-file.")]
        string? source_sheet = null,
        [Description("Target workbook path for copy-to-file and move-to-file.")]
        string? target_file = null,
        [Description(
            "Optional new worksheet name when using copy-to-file. If omitted, the copied sheet keeps its original name.")]
        string? target_sheet_name = null,
        [Description(
            "Optional position control for move, copy-to-file, or move-to-file: insert before this worksheet.")]
        string? before_sheet = null,
        [Description(
            "Optional position control for move, copy-to-file, or move-to-file: insert after this worksheet.")]
        string? after_sheet = null,
        [Description("Return an exact mutation plan without changing Excel. Atomic cross-file actions reject safety options without changing either file.")]
        bool review_only = false,
        [Description("Bound review identifier returned by a prior review-only request.")]
        string? review_id = null,
        [Description("Create a recoverable workbook checkpoint immediately before a mutation.")]
        bool checkpoint = false,
        [Description("Client-generated key for safely retrying a same-workbook mutation.")]
        string? idempotency_key = null,
        CancellationToken cancellationToken = default)
    {
        using var cancellationScope = ExcelToolsBase.PushCancellationToken(cancellationToken);
        using var safetyScope = ExcelToolsBase.PushSafetyOptions(review_only, review_id, checkpoint, idempotency_key);

        return ExcelToolsBase.ExecuteToolAction(
            "worksheet",
            ServiceRegistry.Sheet.ToActionString(action),
            () =>
            {
                // Atomic operations don't require a session.
                if (!ServiceRegistry.Sheet.RequiresSessionForAction(action))
                {
                    return action switch
                    {
                        SheetAction.CopyToFile =>
                            ServiceRegistry.Sheet.RouteAction(
                                action,
                                "",  // No session for atomic operation
                                ExcelToolsBase.ForwardToServiceFunc,
                                sourceFile: source_file,
                                sourceSheet: source_sheet,
                                targetFile: target_file,
                                targetSheetName: target_sheet_name,
                                beforeSheet: before_sheet,
                                afterSheet: after_sheet),
                        SheetAction.MoveToFile =>
                            ServiceRegistry.Sheet.RouteAction(
                                action,
                                "",  // No session for atomic operation
                                ExcelToolsBase.ForwardToServiceFunc,
                                sourceFile: source_file,
                                sourceSheet: source_sheet,
                                targetFile: target_file,
                                beforeSheet: before_sheet,
                                afterSheet: after_sheet),
                        _ => throw new ArgumentException($"Unknown atomic action: {action}"),
                    };
                }

                // Validate session_id for non-atomic operations
                if (string.IsNullOrWhiteSpace(session_id))
                {
                    return JsonSerializer.Serialize(new
                    {
                        success = false,
                        errorMessage = "session_id is required for this action. Use file 'open' action to start a session.",
                        isError = true
                    }, ExcelToolsBase.JsonOptions);
                }

                if (action == SheetAction.Rename)
                {
                    if (string.IsNullOrWhiteSpace(old_name))
                    {
                        throw new ArgumentException("old_name is required for rename action", nameof(old_name));
                    }

                    if (string.IsNullOrWhiteSpace(new_name))
                    {
                        throw new ArgumentException("new_name is required for rename action", nameof(new_name));
                    }
                }

                // Session-based operations
                return action switch
                {
                    SheetAction.List =>
                        ServiceRegistry.Sheet.RouteAction(
                            action,
                            session_id,
                            ExcelToolsBase.ForwardToServiceFunc,
                            filePath: file_path),
                    SheetAction.Create =>
                        ServiceRegistry.Sheet.RouteAction(
                            action,
                            session_id,
                            ExcelToolsBase.ForwardToServiceFunc,
                            sheetName: sheet_name,
                            filePath: file_path),
                    SheetAction.Rename =>
                        ServiceRegistry.Sheet.RouteAction(
                            action,
                            session_id,
                            ExcelToolsBase.ForwardToServiceFunc,
                            oldName: old_name,
                            newName: new_name),
                    SheetAction.Delete =>
                        ServiceRegistry.Sheet.RouteAction(
                            action,
                            session_id,
                            ExcelToolsBase.ForwardToServiceFunc,
                            sheetName: sheet_name),
                    SheetAction.Copy =>
                        ServiceRegistry.Sheet.RouteAction(
                            action,
                            session_id,
                            ExcelToolsBase.ForwardToServiceFunc,
                            sourceName: source_name,
                            targetName: target_name),
                    SheetAction.Move =>
                        ServiceRegistry.Sheet.RouteAction(
                            action,
                            session_id,
                            ExcelToolsBase.ForwardToServiceFunc,
                            sheetName: sheet_name,
                            beforeSheet: before_sheet,
                            afterSheet: after_sheet),
                    _ => throw new ArgumentException($"Unknown action: {action} ({ServiceRegistry.Sheet.ToActionString(action)})", nameof(action))
                };
            });
    }
}
