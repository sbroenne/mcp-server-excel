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
    /// ATOMIC OPERATIONS: copy-to-file and move-to-file don't require a session (open/close automatically).
    /// POSITIONING: Use before OR after (not both) to place sheet relative to another.
    /// Use worksheet_style for tab colors and visibility.
    /// </summary>
    /// <param name="action">The action to perform</param>
    /// <param name="session_id">Session ID from file 'open' action (required for: list, create, rename, delete, move, copy. Not required for: copy-to-file, move-to-file)</param>
    /// <param name="sheet_name">Name of the worksheet (required for: create, rename, delete, move, copy)</param>
    /// <param name="source_name">Name of the source worksheet (required for: copy)</param>
    /// <param name="target_name">Name for the target/copied worksheet</param>
    /// <param name="file_path">Optional file path when batch contains multiple workbooks</param>
    /// <param name="source_file">Full path to the source workbook (required for: copy-to-file, move-to-file)</param>
    /// <param name="source_sheet">Name of the sheet to copy (required for: copy-to-file, move-to-file)</param>
    /// <param name="target_file">Full path to the target workbook (required for: copy-to-file, move-to-file)</param>
    /// <param name="target_sheet_name">Optional: New name for the copied sheet (default: keeps original name)</param>
    /// <param name="before_sheet">Optional: Position before this sheet</param>
    /// <param name="after_sheet">Optional: Position after this sheet</param>
    [McpServerTool(Name = "worksheet", Title = "Worksheet Operations", Destructive = true)]
    [McpMeta("category", "structure")]
    [McpMeta("requiresSession", false)]  // Session is optional - depends on the action
    [Description("Worksheet lifecycle: create, rename, copy, delete, move. ATOMIC OPERATIONS: copy-to-file and move-to-file don't require a session (open/close automatically). POSITIONING: Use before OR after (not both) to place sheet relative to another. Use worksheet_style for tab colors and visibility.")]
    public static string ExcelWorksheet(
        [Description("The action to perform")] SheetAction action,
        [DefaultValue(null)] string? session_id,
        [DefaultValue(null)] string? sheet_name,
        [DefaultValue(null)] string? source_name,
        [DefaultValue(null)] string? target_name,
        [DefaultValue(null)] string? file_path,
        [DefaultValue(null)] string? source_file,
        [DefaultValue(null)] string? source_sheet,
        [DefaultValue(null)] string? target_file,
        [DefaultValue(null)] string? target_sheet_name,
        [DefaultValue(null)] string? before_sheet,
        [DefaultValue(null)] string? after_sheet)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "worksheet",
            ServiceRegistry.Sheet.ToActionString(action),
            () =>
            {
                // Atomic operations don't require a session
                if (action == SheetAction.CopyToFile || action == SheetAction.MoveToFile)
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
                            oldName: sheet_name,
                            newName: target_name),
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
