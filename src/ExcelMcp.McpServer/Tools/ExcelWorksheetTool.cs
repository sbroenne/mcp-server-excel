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
    /// Worksheet lifecycle. CopyToFile/MoveToFile are atomic (no session).
    /// Other actions need sid. Position: before OR after (not both).
    /// Related: excel_worksheet_style (colors/visibility)
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="sid">Session ID</param>
    /// <param name="src">Source file (copy-to-file/move-to-file)</param>
    /// <param name="sn">Sheet name</param>
    /// <param name="tgt">Target file</param>
    /// <param name="tn">New name (rename/copy)</param>
    /// <param name="before">Position before sheet</param>
    /// <param name="after">Position after sheet</param>
    [McpServerTool(Name = "excel_worksheet", Title = "Excel Worksheet Operations")]
    [McpMeta("category", "structure")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelWorksheet(
        WorksheetAction action,
        [DefaultValue(null)] string? sid,
        [DefaultValue(null)] string? src,
        [DefaultValue(null)] string? sn,
        [DefaultValue(null)] string? tgt,
        [DefaultValue(null)] string? tn,
        [DefaultValue(null)] string? before,
        [DefaultValue(null)] string? after)
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
                    WorksheetAction.CopyToFile => CopyToFileHandler(sheetCommands, src, sn, tgt, tn, before, after),
                    WorksheetAction.MoveToFile => MoveToFileHandler(sheetCommands, src, sn, tgt, before, after),

                    // Session-based operations
                    WorksheetAction.List => ListAsync(sheetCommands, sid!),
                    WorksheetAction.Create => CreateAsync(sheetCommands, sid!, sn),
                    WorksheetAction.Delete => DeleteAsync(sheetCommands, sid!, sn),
                    WorksheetAction.Rename => RenameAsync(sheetCommands, sid!, sn, tn),
                    WorksheetAction.Copy => CopyAsync(sheetCommands, sid!, sn, tn),
                    WorksheetAction.Move => MoveAsync(sheetCommands, sid!, sn, before, after),
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

