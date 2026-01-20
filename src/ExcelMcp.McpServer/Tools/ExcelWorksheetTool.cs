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
        WorksheetAction action,
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
            action.ToActionString(),
            () =>
            {
                var sheetCommands = new SheetCommands();

                // Expression switch pattern for audit compliance
                return action switch
                {
                    // Atomic cross-file operations (no session required)
                    WorksheetAction.CopyToFile => CopyToFileHandler(sheetCommands, sourceFile, sheetName, targetFile, targetName, beforeSheet, afterSheet),
                    WorksheetAction.MoveToFile => MoveToFileHandler(sheetCommands, sourceFile, sheetName, targetFile, beforeSheet, afterSheet),

                    // Session-based operations
                    WorksheetAction.List => ListAsync(sheetCommands, sessionId!),
                    WorksheetAction.Create => CreateAsync(sheetCommands, sessionId!, sheetName),
                    WorksheetAction.Delete => DeleteAsync(sheetCommands, sessionId!, sheetName),
                    WorksheetAction.Rename => RenameAsync(sheetCommands, sessionId!, sheetName, targetName),
                    WorksheetAction.Copy => CopyAsync(sheetCommands, sessionId!, sheetName, targetName),
                    WorksheetAction.Move => MoveAsync(sheetCommands, sessionId!, sheetName, beforeSheet, afterSheet),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    // === PRIVATE HELPER METHODS ===

    private static string ListAsync(
        SheetCommands sheetCommands,
        string sessionId)
    {
        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.List(batch));

        return JsonSerializer.Serialize(new
        {
            success = result.Success,
            worksheets = result.Worksheets,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for create action", nameof(sheetName));

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Create(batch, sheetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' created successfully."
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

    private static string RenameAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ArgumentException("sheetName and targetName are required for rename action", "sheetName,targetName");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Rename(batch, sheetName, targetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' renamed to '{targetName}' successfully."
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

    private static string CopyAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ArgumentException("sheetName and targetName are required for copy action", "sheetName,targetName");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Copy(batch, sheetName, targetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' copied to '{targetName}' successfully."
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

    private static string DeleteAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for delete action", nameof(sheetName));

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Delete(batch, sheetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' deleted successfully."
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

    private static string MoveAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? beforeSheet,
        string? afterSheet)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for move action", nameof(sheetName));

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Move(batch, sheetName, beforeSheet, afterSheet);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' moved successfully."
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

    // === ATOMIC CROSS-FILE OPERATIONS ===

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

