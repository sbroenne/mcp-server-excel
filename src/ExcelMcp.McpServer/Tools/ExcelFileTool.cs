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
    /// File and session management for Excel automation.
    ///
    /// WORKFLOW: open → use sessionId with other tools → close (save=true to persist changes).
    ///
    /// IMPORTANT: Before closing, check 'list' action - wait for canClose=true (no active operations).
    /// If showExcel=true was used, confirm with user before closing visible Excel windows.
    /// </summary>
    /// <param name="action">The file operation to perform</param>
    /// <param name="excelPath">Full path to Excel file (.xlsx or .xlsm). Required for: open, create-empty, test</param>
    /// <param name="sessionId">Session ID returned from 'open' action. Required for: close. Used by all other tools.</param>
    /// <param name="save">Whether to save changes when closing. Default: false (discard changes)</param>
    /// <param name="showExcel">Whether to make Excel window visible. Default: false (hidden automation)</param>
    [McpServerTool(Name = "excel_file", Title = "Excel File Operations")]
    [McpMeta("category", "session")]
    [McpMeta("requiresSession", false)]
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

