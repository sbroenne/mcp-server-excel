using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel worksheet lifecycle (create, rename, copy, delete, move).
/// Use excel_worksheet_style for tab colors and visibility.
/// </summary>
[McpServerToolType]
public static partial class ExcelWorksheetTool
{
    /// <summary>
    /// Worksheet lifecycle management: create, rename, copy, delete, move sheets.
    ///
    /// ATOMIC OPERATIONS: 'copy-to-file' and 'move-to-file' don't require a session - they open/close files automatically.
    /// SESSION OPERATIONS: All other actions require a sessionId from excel_file 'open'.
    ///
    /// POSITIONING: Use 'before' OR 'after' (not both) to position the sheet relative to another.
    ///
    /// Related: Use excel_worksheet_style for tab colors and visibility settings.
    /// </summary>
    /// <param name="action">The worksheet operation to perform</param>
    /// <param name="sessionId">Session ID from excel_file 'open'. Required for: list, create, delete, rename, copy, move</param>
    /// <param name="sourceFile">Full path to source Excel file. Required for: copy-to-file, move-to-file</param>
    /// <param name="sheetName">Name of the worksheet to operate on. Required for: create, delete, rename, copy, move, copy-to-file, move-to-file</param>
    /// <param name="targetFile">Full path to destination Excel file. Required for: copy-to-file, move-to-file</param>
    /// <param name="targetName">New name for the worksheet. Required for: rename. Optional for: copy, copy-to-file (defaults to sheetName)</param>
    /// <param name="beforeSheet">Position the sheet before this sheet name. Optional for: move, copy-to-file, move-to-file</param>
    /// <param name="afterSheet">Position the sheet after this sheet name. Optional for: move, copy-to-file, move-to-file</param>
    [McpServerTool(Name = "excel_worksheet", Title = "Excel Worksheet Operations", Destructive = true)]
    [McpMeta("category", "structure")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelWorksheet(
        SheetAction action,
        [DefaultValue(null)] string? sessionId,
        [DefaultValue(null)] string? sourceFile,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? targetFile,
        [DefaultValue(null)] string? targetName,
        [DefaultValue(null)] string? beforeSheet,
        [DefaultValue(null)] string? afterSheet)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_worksheet",
            ServiceRegistry.Sheet.ToActionString(action),
            () =>
            {
                // Expression switch pattern for audit compliance
                return action switch
                {
                    // Atomic cross-file operations (no session required - handled locally)
                    SheetAction.CopyToFile => CopyToFileHandler(new SheetCommands(), sourceFile, sheetName, targetFile, targetName, beforeSheet, afterSheet),
                    SheetAction.MoveToFile => MoveToFileHandler(new SheetCommands(), sourceFile, sheetName, targetFile, beforeSheet, afterSheet),

                    // Session-based operations - forward to service
                    SheetAction.List => ExcelToolsBase.ForwardToService("sheet.list", sessionId!, null),
                    SheetAction.Create => ForwardCreate(sessionId!, sheetName),
                    SheetAction.Delete => ForwardDelete(sessionId!, sheetName),
                    SheetAction.Rename => ForwardRename(sessionId!, sheetName, targetName),
                    SheetAction.Copy => ForwardCopy(sessionId!, sheetName, targetName),
                    SheetAction.Move => ForwardMove(sessionId!, sheetName, beforeSheet, afterSheet),
                    _ => throw new ArgumentException($"Unknown action: {action} ({ServiceRegistry.Sheet.ToActionString(action)})", nameof(action))
                };
            });
    }

    // === SERVICE FORWARDING METHODS ===

    private static string ForwardCreate(string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for create action", nameof(sheetName));

        return ExcelToolsBase.ForwardToService("sheet.create", sessionId, new { sheetName });
    }

    private static string ForwardDelete(string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for delete action", nameof(sheetName));

        return ExcelToolsBase.ForwardToService("sheet.delete", sessionId, new { sheetName });
    }

    private static string ForwardRename(string sessionId, string? sheetName, string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ArgumentException("sheetName and targetName are required for rename action", "sheetName,targetName");

        return ExcelToolsBase.ForwardToService("sheet.rename", sessionId, new { sheetName, newName = targetName });
    }

    private static string ForwardCopy(string sessionId, string? sheetName, string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for copy action", nameof(sheetName));

        return ExcelToolsBase.ForwardToService("sheet.copy", sessionId, new { sourceSheet = sheetName, targetSheet = targetName ?? sheetName });
    }

    private static string ForwardMove(string sessionId, string? sheetName, string? beforeSheet, string? afterSheet)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for move action", nameof(sheetName));

        return ExcelToolsBase.ForwardToService("sheet.move", sessionId, new { sheetName, beforeSheet, afterSheet });
    }

    // === ATOMIC CROSS-FILE OPERATIONS ===
    // These don't use sessions - they open/close files atomically

    private static string CopyToFileHandler(
        SheetCommands sheetCommands,
        string? sourceFile,
        string? sheetName,
        string? targetFile,
        string? targetName,
        string? beforeSheet,
        string? afterSheet)
    {
        if (string.IsNullOrEmpty(sourceFile))
            throw new ArgumentException("sourceFile is required for copy-to-file action", nameof(sourceFile));

        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for copy-to-file action", nameof(sheetName));

        if (string.IsNullOrEmpty(targetFile))
            throw new ArgumentException("targetFile is required for copy-to-file action", nameof(targetFile));

        try
        {
            sheetCommands.CopyToFile(sourceFile, sheetName, targetFile, targetName, beforeSheet, afterSheet);

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' copied from '{Path.GetFileName(sourceFile)}' to '{Path.GetFileName(targetFile)}' successfully.",
                sourceFile,
                targetFile,
                sheetName,
                targetSheetName = targetName ?? sheetName
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string MoveToFileHandler(
        SheetCommands sheetCommands,
        string? sourceFile,
        string? sheetName,
        string? targetFile,
        string? beforeSheet,
        string? afterSheet)
    {
        if (string.IsNullOrEmpty(sourceFile))
            throw new ArgumentException("sourceFile is required for move-to-file action", nameof(sourceFile));

        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for move-to-file action", nameof(sheetName));

        if (string.IsNullOrEmpty(targetFile))
            throw new ArgumentException("targetFile is required for move-to-file action", nameof(targetFile));

        try
        {
            sheetCommands.MoveToFile(sourceFile, sheetName, targetFile, beforeSheet, afterSheet);

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' moved from '{Path.GetFileName(sourceFile)}' to '{Path.GetFileName(targetFile)}' successfully.",
                sourceFile,
                targetFile,
                sheetName
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }
}





