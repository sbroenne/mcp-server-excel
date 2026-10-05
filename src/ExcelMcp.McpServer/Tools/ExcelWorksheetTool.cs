using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel worksheet management tool for MCP server.
/// Handles session-based changes (create, rename, delete, move, copy)
/// and atomic cross-file operations (copy-to-file, move-to-file).
/// Use worksheet_read to list worksheets.
/// </summary>
[McpServerToolType]
public static partial class ExcelWorksheetTool
{
    /// <summary>
    /// Worksheet lifecycle: create, rename, copy, delete, move.
    /// delete removes the sheet and all its contents and may break dependent references; no tool-level undo.
    /// move-to-file removes the source sheet and saves both files; no tool-level undo.
    /// RENAME: Use old_name + new_name.
    /// ATOMIC OPERATIONS: copy-to-file and move-to-file don't require a session (open/close automatically).
    /// POSITIONING: Use before_sheet or after_sheet (not both) to place a sheet relative to another.
    /// Use worksheet_style for tab colors, visibility, and protection.
    /// </summary>
    /// <param name="action">The worksheet change to perform</param>
    /// <param name="workbook_session_id">Session ID from file 'open' or 'create' (required for same-workbook changes; not used for copy-to-file or move-to-file)</param>
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
    [McpServerTool(Name = "worksheet", Title = "Worksheet Operations", Destructive = true,
        UseStructuredContent = true, OutputSchemaType = typeof(WorksheetToolOutputSchema))]
    [McpMeta("category", "structure")]
    [McpMeta("requiresSession", false)]  // Session is optional - depends on the action
    [Description("Worksheet changes: create, rename, copy, delete, move. DELETE HAS NO TOOL-LEVEL UNDO: removes all sheet contents and may break dependent references; check the intended sheet and its dependencies. MOVE-TO-FILE HAS NO TOOL-LEVEL UNDO: removes the source sheet and saves both files. Closing another session without saving cannot reverse that transfer. Rename uses old_name and new_name. Cross-file copy-to-file and move-to-file open, save, and close automatically without a session. Position with before_sheet or after_sheet, not both. Use worksheet_style for tab colors, visibility, and protection.")]
    public static Task<CallToolResult> ExcelWorksheet(
        [Description("The worksheet change to perform")] WorksheetWriteAction action,
        ServiceBridge.ServiceBridge bridge,
        [Description(
            "Session ID from file 'open' or 'create'. Required for same-workbook changes: create, rename, delete, move, and copy. Not used by copy-to-file or move-to-file.")]
        string? workbook_session_id = null,
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
        CancellationToken cancellationToken = default)
    {
        var serviceAction = action switch
        {
            WorksheetWriteAction.Create => SheetAction.Create,
            WorksheetWriteAction.Rename => SheetAction.Rename,
            WorksheetWriteAction.Copy => SheetAction.Copy,
            WorksheetWriteAction.Delete => SheetAction.Delete,
            WorksheetWriteAction.Move => SheetAction.Move,
            WorksheetWriteAction.CopyToFile => SheetAction.CopyToFile,
            WorksheetWriteAction.MoveToFile => SheetAction.MoveToFile,
            _ => throw new ArgumentOutOfRangeException(nameof(action))
        };
        return ExcelToolsBase.ExecuteToolActionAsync(
            "worksheet",
            ServiceRegistry.Sheet.ToActionString(serviceAction),
            async () =>
            {
                // Atomic operations don't require a session
                if (serviceAction == SheetAction.CopyToFile || serviceAction == SheetAction.MoveToFile)
                {
                    return await (serviceAction switch
                    {
                        SheetAction.CopyToFile =>
                            ServiceRegistry.Sheet.RouteAction(
                                serviceAction,
                                "",  // No session for atomic operation
                                (command, id, args) => ExcelToolsBase.ForwardToServiceAsync(bridge, command, id, args, cancellationToken),
                                sourceFile: source_file,
                                sourceSheet: source_sheet,
                                targetFile: target_file,
                                targetSheetName: target_sheet_name,
                                beforeSheet: before_sheet,
                                afterSheet: after_sheet),
                        SheetAction.MoveToFile =>
                            ServiceRegistry.Sheet.RouteAction(
                                serviceAction,
                                "",  // No session for atomic operation
                                (command, id, args) => ExcelToolsBase.ForwardToServiceAsync(bridge, command, id, args, cancellationToken),
                                sourceFile: source_file,
                                sourceSheet: source_sheet,
                                targetFile: target_file,
                                beforeSheet: before_sheet,
                                afterSheet: after_sheet),
                        _ => throw new ArgumentException($"Unknown atomic action: {serviceAction}"),
                    });
                }

                // Validate the session input for non-atomic operations.
                if (string.IsNullOrWhiteSpace(workbook_session_id))
                {
                    return JsonSerializer.Serialize(new
                    {
                        success = false,
                        errorMessage = "workbook_session_id is required for this action. Use file 'open' action to start a session.",
                        errorCategory = "InvalidInput",
                        isError = true
                    }, ExcelToolsBase.JsonOptions);
                }

                if (serviceAction == SheetAction.Rename)
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
                return await (serviceAction switch
                {
                    SheetAction.Create =>
                        ServiceRegistry.Sheet.RouteAction(
                            serviceAction,
                            workbook_session_id,
                            (command, id, args) => ExcelToolsBase.ForwardToServiceAsync(bridge, command, id, args, cancellationToken),
                            sheetName: sheet_name,
                            filePath: file_path),
                    SheetAction.Rename =>
                        ServiceRegistry.Sheet.RouteAction(
                            serviceAction,
                            workbook_session_id,
                            (command, id, args) => ExcelToolsBase.ForwardToServiceAsync(bridge, command, id, args, cancellationToken),
                            oldName: old_name,
                            newName: new_name),
                    SheetAction.Delete =>
                        ServiceRegistry.Sheet.RouteAction(
                            serviceAction,
                            workbook_session_id,
                            (command, id, args) => ExcelToolsBase.ForwardToServiceAsync(bridge, command, id, args, cancellationToken),
                            sheetName: sheet_name),
                    SheetAction.Copy =>
                        ServiceRegistry.Sheet.RouteAction(
                            serviceAction,
                            workbook_session_id,
                            (command, id, args) => ExcelToolsBase.ForwardToServiceAsync(bridge, command, id, args, cancellationToken),
                            sourceName: source_name,
                            targetName: target_name),
                    SheetAction.Move =>
                        ServiceRegistry.Sheet.RouteAction(
                            serviceAction,
                            workbook_session_id,
                            (command, id, args) => ExcelToolsBase.ForwardToServiceAsync(bridge, command, id, args, cancellationToken),
                            sheetName: sheet_name,
                            beforeSheet: before_sheet,
                            afterSheet: after_sheet),
                    _ => throw new ArgumentException($"Unknown action: {serviceAction} ({ServiceRegistry.Sheet.ToActionString(serviceAction)})", nameof(action))
                });
            }, cancellationToken);
    }

    /// <summary>List worksheets in a workbook session.</summary>
    /// <param name="action">List worksheets.</param>
    /// <param name="workbook_session_id">Session ID returned by file open/create or file_read list.</param>
    /// <param name="file_path">Optional workbook path when the session has multiple open workbooks.</param>
    [McpServerTool(Name = "worksheet_read", Title = "Read-Only Worksheet Operations",
        ReadOnly = true, Destructive = false, UseStructuredContent = true,
        OutputSchemaType = typeof(WorksheetToolOutputSchema))]
    [McpMeta("category", "structure")]
    [McpMeta("requiresSession", true)]
    [Description("List worksheets in a workbook session.")]
    public static Task<CallToolResult> ExcelWorksheetRead(
        [Description("The read-only action to perform")] WorksheetReadAction action,
        ServiceBridge.ServiceBridge bridge,
        [Description("Session ID returned by file open/create or file_read list.")]
        string workbook_session_id,
        [Description("Optional workbook path when the session has multiple open workbooks.")]
        string? file_path = null,
        CancellationToken cancellationToken = default) =>
        ExcelToolsBase.ExecuteToolActionAsync(
            "worksheet_read",
            "list",
            () => ServiceRegistry.Sheet.RouteAction(
                action switch
                {
                    WorksheetReadAction.List => SheetAction.List,
                    _ => throw new ArgumentOutOfRangeException(nameof(action))
                },
                workbook_session_id,
                (command, id, args) => ExcelToolsBase.ForwardToServiceAsync(bridge, command, id, args, cancellationToken),
                filePath: file_path),
            cancellationToken);
}
