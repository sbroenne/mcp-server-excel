using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel file management tool for MCP server.
/// </summary>
[McpServerToolType]
public static partial class ExcelFileTool
{
    /// <summary>
    /// Manage Excel files and sessions.
    ///
    /// SESSION VERIFICATION (use instead of calling other tools):
    /// - LIST - Returns all active sessions with status (activeOperations, canClose, isExcelVisible)
    /// - Use to check if operations are still running before attempting close
    /// - Use to recover sessionId if lost
    ///
    /// SESSION LIFECYCLE:
    /// 1. OPEN - Start session, get sessionId
    /// 2. OPERATE - Use sessionId with other tools
    /// 3. CLOSE - End session (use save:true parameter to persist changes)
    ///
    /// IMPORTANT: NO 'SAVE' ACTION - Use action='close' with save:true to persist changes
    ///
    /// OPERATION TRACKING (automatic protection):
    /// - Server tracks active operations per session
    /// - Close is BLOCKED if operations are still running
    /// - Use LIST to check activeOperations count before closing
    /// - Wait for canClose=true before closing
    ///
    /// WHEN showExcel=true - ASK BEFORE CLOSING:
    /// - If Excel is visible (isExcelVisible=true in LIST), the user is actively watching
    /// - ALWAYS ask user before closing: "Would you like me to save and close, or keep it open?"
    /// - User may want to inspect results, make manual changes, or continue working
    /// - Do NOT auto-close visible Excel sessions
    ///
    /// WORKFLOWS:
    /// - Verify session: list → check sessionId exists
    /// - Check before close: list → verify canClose=true → close
    /// - Persist changes: open → operations(sessionId) → close(save: true)
    /// - Discard changes: open → operations(sessionId) → close(save: false)
    ///
    /// FILE FORMATS:
    /// - .xlsx: Standard Excel workbook
    /// - .xlsm: Macro-enabled workbook
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="excelPath">Excel file path (.xlsx or .xlsm) - required for open/create-empty, not used for close</param>
    /// <param name="sessionId">Session ID from 'open' action - required for close</param>
    /// <param name="save">Save changes before closing (default: false)</param>
    /// <param name="showExcel">Show Excel window during operations (default: false). Set true so user can watch changes in real-time.</param>
    [McpServerTool(Name = "excel_file")]
    [McpMeta("category", "session")]
    public static partial string ExcelFile(
        FileAction action,
        [DefaultValue(null)] string? excelPath,
        [DefaultValue(null)] string? sessionId,
        [DefaultValue(false)] bool save,
        [DefaultValue(false)] bool showExcel)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_file",
            action.ToActionString(),
            excelPath,
            () =>
            {
                var fileCommands = new FileCommands();

                // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
                return action switch
                {
                    FileAction.List => ListSessions(),
                    FileAction.Open => OpenSessionAsync(excelPath!, showExcel),
                    FileAction.Close => CloseSessionAsync(sessionId!, save),
                    FileAction.CreateEmpty => CreateEmptyFileAsync(fileCommands, excelPath!,
                        excelPath!.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)),
                    FileAction.CloseWorkbook => CloseWorkbook(excelPath!),
                    FileAction.Test => TestFileAsync(fileCommands, excelPath!),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    /// <summary>
    /// Opens an Excel file and creates a new session.
    /// Returns sessionId that must be used for all subsequent operations.
    /// </summary>
    private static string OpenSessionAsync(string excelPath, bool showExcel)
    {
        if (string.IsNullOrWhiteSpace(excelPath))
        {
            throw new ArgumentException("excelPath is required for 'open' action", nameof(excelPath));
        }

        if (!File.Exists(excelPath))
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"File not found: {excelPath}",
                filePath = excelPath,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }

        try
        {
            string sessionId = ExcelToolsBase.GetSessionManager().CreateSession(excelPath, showExcel);

            return JsonSerializer.Serialize(new
            {
                success = true,
                sessionId,
                filePath = excelPath,
                showExcel
            }, ExcelToolsBase.JsonOptions);
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("already open"))
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"Cannot open '{excelPath}': {ex.Message}",
                filePath = excelPath,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"Cannot open '{excelPath}': {ex.Message}",
                filePath = excelPath,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    /// <summary>
    /// Closes an active session with optional save.
    /// By default, saves changes before closing to prevent data loss.
    /// Set save=false to discard changes.
    /// </summary>
    private static string CloseSessionAsync(string sessionId, bool save)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            throw new ArgumentException("sessionId is required for 'close' action", nameof(sessionId));
        }

        var sessionManager = ExcelToolsBase.GetSessionManager();

        // Validate before closing - check for running operations
        var validation = sessionManager.ValidateClose(sessionId);

        if (!validation.SessionExists)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                sessionId,
                errorMessage = validation.BlockingReason ?? $"Session '{sessionId}' not found",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }

        if (validation.ActiveOperationCount > 0)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                sessionId,
                errorMessage = validation.BlockingReason,
                activeOperations = validation.ActiveOperationCount,
                isExcelVisible = validation.IsExcelVisible,
                suggestedAction = "Wait for all operations to complete before closing the session.",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }

        try
        {
            bool success = sessionManager.CloseSession(sessionId, save);

            if (success)
            {
                return JsonSerializer.Serialize(new
                {
                    success = true,
                    sessionId,
                    saved = save
                }, ExcelToolsBase.JsonOptions);
            }

            return JsonSerializer.Serialize(new
            {
                success = false,
                sessionId,
                errorMessage = $"Session '{sessionId}' not found",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                sessionId,
                errorMessage = $"Cannot close session '{sessionId}': {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    /// <summary>
    /// Creates a new empty Excel file (.xlsx or .xlsm based on macroEnabled flag).
    /// LLM Pattern: Use this when you need a fresh Excel workbook for automation.
    /// </summary>
    private static string CreateEmptyFileAsync(FileCommands fileCommands, string excelPath, bool macroEnabled)
    {
        var extension = macroEnabled ? ".xlsm" : ".xlsx";
        if (!excelPath.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
        {
            excelPath = Path.ChangeExtension(excelPath, extension);
        }

        fileCommands.CreateEmpty(excelPath, overwriteIfExists: false);

        return JsonSerializer.Serialize(new
        {
            success = true,
            filePath = excelPath,
            macroEnabled,
            message = "Excel workbook created successfully"
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Closes the workbook (no-op with new single-instance architecture).
    /// LLM Pattern: This action is kept for backward compatibility but does nothing.
    /// With single-instance sessions, workbooks are automatically closed after each operation.
    /// </summary>
    private static string CloseWorkbook(string excelPath)
    {
        return JsonSerializer.Serialize(new
        {
            success = true,
            filePath = excelPath
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Lists all active sessions with status info. Lightweight operation - no Excel COM calls.
    /// LLM Pattern: Use this to verify sessions and check for running operations before closing.
    /// </summary>
    private static string ListSessions()
    {
        var sessionManager = ExcelToolsBase.GetSessionManager();
        var sessions = sessionManager.GetActiveSessions();

        var sessionList = sessions.Select(s => new
        {
            sessionId = s.SessionId,
            filePath = s.FilePath,
            activeOperations = sessionManager.GetActiveOperationCount(s.SessionId),
            isExcelVisible = sessionManager.IsExcelVisible(s.SessionId),
            canClose = sessionManager.GetActiveOperationCount(s.SessionId) == 0
        }).ToList();

        return JsonSerializer.Serialize(new
        {
            success = true,
            sessions = sessionList,
            count = sessionList.Count
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Tests if an Excel file exists and is valid without opening it via Excel COM.
    /// LLM Pattern: Use this for discovery/connectivity testing before running operations.
    /// </summary>
    private static string TestFileAsync(FileCommands fileCommands, string excelPath)
    {
        if (string.IsNullOrWhiteSpace(excelPath))
        {
            throw new ArgumentException("excelPath is required for 'test' action", nameof(excelPath));
        }

        var info = fileCommands.Test(excelPath);

        return JsonSerializer.Serialize(new
        {
            success = info.IsValid,
            filePath = info.FilePath,
            exists = info.Exists,
            isValid = info.IsValid,
            extension = info.Extension,
            size = info.Size,
            lastModified = info.LastModified,
            message = info.Message,
            isError = info.IsValid ? (bool?)null : true
        }, ExcelToolsBase.JsonOptions);
    }
}

