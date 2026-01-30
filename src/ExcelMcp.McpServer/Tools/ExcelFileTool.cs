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
    /// NEW FILES: Use create-and-open for single optimized operation (50% faster than open+create separately).
    ///
    /// SESSION REUSE: Call 'list' first to check for existing sessions.
    /// If file is already open, reuse existing sessionId instead of opening again.
    ///
    /// IMPORTANT: Before closing, check 'list' action - wait for canClose=true (no active operations).
    /// If showExcel=true was used, confirm with user before closing visible Excel windows.
    ///
    /// TIMEOUT: Each operation has a 5-min default timeout. Use timeoutSeconds to customize
    /// for long-running operations (data refresh, large queries). Operations timing out
    /// trigger aggressive cleanup and may leave Excel in inconsistent state.
    /// </summary>
    /// <param name="action">The file operation to perform</param>
    /// <param name="excelPath">Full Windows path to Excel file (.xlsx or .xlsm). ASK USER for the path - do not guess or use placeholder usernames. Required for: open, create-and-open, test</param>
    /// <param name="sessionId">Session ID returned from 'open' or 'create-and-open'. Required for: close. Used by all other tools.</param>
    /// <param name="save">Whether to save changes when closing. Default: false (discard changes)</param>
    /// <param name="showExcel">Whether to make Excel window visible. Default: false (hidden automation)</param>
    /// <param name="timeoutSeconds">Maximum time in seconds for any operation in this session. Default: 300 (5 min). Range: 10-3600. Used for: open, create-and-open</param>
    [McpServerTool(Name = "excel_file", Title = "Excel File Operations", Destructive = true)]
    [McpMeta("category", "session")]
    [McpMeta("requiresSession", false)]
    public static partial string ExcelFile(
        FileAction action,
        [DefaultValue(null)] string? excelPath,
        [DefaultValue(null)] string? sessionId,
        [DefaultValue(false)] bool save,
        [DefaultValue(false)] bool showExcel,
        [DefaultValue(300)] int timeoutSeconds)
    {
        // Validate timeout range
        if (timeoutSeconds < 10 || timeoutSeconds > 3600)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"timeoutSeconds must be between 10 and 3600 seconds, got {timeoutSeconds}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }

        var timeout = TimeSpan.FromSeconds(timeoutSeconds);

        return ExcelToolsBase.ExecuteToolAction(
            "excel_file",
            action.ToActionString(),
            excelPath,
            () =>
            {
                // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
                return action switch
                {
                    FileAction.List => ListSessions(),
                    FileAction.Open => OpenSessionAsync(excelPath!, showExcel, timeout),
                    FileAction.Close => CloseSessionAsync(sessionId!, save),
                    FileAction.CreateAndOpen => CreateAndOpenSessionAsync(excelPath!, showExcel, timeout),
                    FileAction.CloseWorkbook => CloseWorkbook(excelPath!),
                    FileAction.Test => TestFileAsync(excelPath!),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    /// <summary>
    /// Opens an Excel file and creates a new session.
    /// Returns sessionId that must be used for all subsequent operations.
    /// </summary>
    private static string OpenSessionAsync(string excelPath, bool showExcel, TimeSpan timeout)
    {
        if (string.IsNullOrWhiteSpace(excelPath))
        {
            throw new ArgumentException("excelPath is required for 'open' action", nameof(excelPath));
        }

        // Validate Windows path format before any file operations
        var pathError = ExcelToolsBase.ValidateWindowsPath(excelPath);
        if (pathError != null)
        {
            return pathError;
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
            string sessionId = ExcelToolsBase.GetSessionManager().CreateSession(excelPath, showExcel, timeout);

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
    /// Creates a new empty Excel file AND opens a session in one operation.
    /// Returns sessionId that must be used for all subsequent operations.
    /// Directory must exist - will not be created automatically.
    /// </summary>
    private static string CreateAndOpenSessionAsync(string excelPath, bool showExcel, TimeSpan timeout)
    {
        if (string.IsNullOrWhiteSpace(excelPath))
        {
            throw new ArgumentException("excelPath is required for 'create-and-open' action", nameof(excelPath));
        }

        // Validate Windows path format before any file operations
        var pathError = ExcelToolsBase.ValidateWindowsPath(excelPath);
        if (pathError != null)
        {
            return pathError;
        }

        // Determine if macro-enabled from extension
        bool macroEnabled = excelPath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase);
        var extension = macroEnabled ? ".xlsm" : ".xlsx";
        if (!excelPath.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
        {
            excelPath = Path.ChangeExtension(excelPath, extension);
        }

        try
        {
            // Use the combined create+open which starts Excel only once
            string sessionId = ExcelToolsBase.GetSessionManager().CreateSessionForNewFile(excelPath, showExcel, timeout);

            return JsonSerializer.Serialize(new
            {
                success = true,
                sessionId,
                filePath = excelPath,
                macroEnabled,
                showExcel,
                message = "Excel workbook created and opened successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("already exists"))
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"Cannot create '{excelPath}': {ex.Message}",
                filePath = excelPath,
                suggestedAction = "Use 'open' action to open existing files, or choose a different file path.",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"Cannot create '{excelPath}': {ex.Message}",
                filePath = excelPath,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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
    private static string TestFileAsync(string excelPath)
    {
        if (string.IsNullOrWhiteSpace(excelPath))
        {
            throw new ArgumentException("excelPath is required for 'test' action", nameof(excelPath));
        }

        // Validate Windows path format before any file operations
        var pathError = ExcelToolsBase.ValidateWindowsPath(excelPath);
        if (pathError != null)
        {
            return pathError;
        }

        var fileCommands = new FileCommands();
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

