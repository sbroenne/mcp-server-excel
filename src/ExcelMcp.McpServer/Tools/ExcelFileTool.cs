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
    /// NEW FILES: Use 'create' action to create file AND start session in one call.
    ///
    /// SESSION REUSE: Call 'list' first to check for existing sessions.
    /// If file is already open, reuse existing sessionId instead of opening again.
    ///
    /// IMPORTANT: Before closing, check 'list' action - wait for canClose=true (no active operations).
    /// If show=true was used, confirm with user before closing visible Excel windows.
    ///
    /// TIMEOUT: Each operation has a 5-min default timeout. Use timeoutSeconds to customize
    /// for long-running operations (data refresh, large queries). Operations timing out
    /// trigger aggressive cleanup and may leave Excel in inconsistent state.
    /// </summary>
    /// <param name="action">The file operation to perform</param>
    /// <param name="path">Full Windows path to Excel file (.xlsx or .xlsm). ASK USER for the path - do not guess or use placeholder usernames. Required for: open, create, test</param>
    /// <param name="session_id">Session ID returned from 'open' or 'create'. Required for: close. Used by all other tools.</param>
    /// <param name="save">Whether to save changes when closing. Default: false (discard changes)</param>
    /// <param name="show">Whether to make Excel window visible. Default: false (hidden automation)</param>
    /// <param name="timeout_seconds">Maximum time in seconds for any operation in this session. Default: 300 (5 min). Range: 10-3600. Used for: open, create</param>
    [McpServerTool(Name = "excel_file", Title = "Excel File Operations", Destructive = true)]
    [McpMeta("category", "session")]
    [McpMeta("requiresSession", false)]
    public static partial string ExcelFile(
        FileAction action,
        [DefaultValue(null)] string? path,
        [DefaultValue(null)] string? session_id,
        [DefaultValue(false)] bool save,
        [DefaultValue(false)] bool show,
        [DefaultValue(300)] int timeout_seconds)
    {
        // Validate timeout range
        if (timeout_seconds < 10 || timeout_seconds > 3600)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"timeout_seconds must be between 10 and 3600 seconds, got {timeout_seconds}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }

        var timeout = TimeSpan.FromSeconds(timeout_seconds);

        return ExcelToolsBase.ExecuteToolAction(
            "excel_file",
            action.ToActionString(),
            path,
            () =>
            {
                // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
                return action switch
                {
                    FileAction.List => ListSessions(),
                    FileAction.Open => OpenSessionAsync(path!, show, timeout),
                    FileAction.Close => CloseSessionAsync(session_id!, save),
                    FileAction.Create => CreateSessionAsync(path!, show, timeout),
                    FileAction.CloseWorkbook => CloseWorkbook(path!),
                    FileAction.Test => TestFileAsync(path!),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    /// <summary>
    /// Opens an Excel file and creates a new session via the ExcelMCP Service.
    /// Returns sessionId that must be used for all subsequent operations.
    /// </summary>
    private static string OpenSessionAsync(string path, bool show, TimeSpan timeout)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("path is required for 'open' action", nameof(path));
        }

        // Validate Windows path format before any file operations
        var pathError = ExcelToolsBase.ValidateWindowsPath(path);
        if (pathError != null)
        {
            return pathError;
        }

        if (!File.Exists(path))
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"File not found: {path}",
                filePath = path,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }

        var timeoutSeconds = (int)timeout.TotalSeconds;
        var response = ExcelToolsBase.SendToServiceAsync(
            "session.open",
            null,
            new { filePath = path, show = show, timeoutSeconds },
            timeoutSeconds
        ).GetAwaiter().GetResult();

        if (!response.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = response.ErrorMessage ?? "Failed to open session",
                filePath = path,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }

        return response.Result ?? JsonSerializer.Serialize(new
        {
            success = true,
            filePath = path,
            show
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Closes an active session via the ExcelMCP Service with optional save.
    /// By default, saves changes before closing to prevent data loss.
    /// Set save=false to discard changes.
    /// </summary>
    private static string CloseSessionAsync(string sessionId, bool save)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            throw new ArgumentException("sessionId is required for 'close' action", nameof(sessionId));
        }

        var response = ExcelToolsBase.SendToServiceAsync(
            "session.close",
            sessionId,
            new { save }
        ).GetAwaiter().GetResult();

        if (!response.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                sessionId,
                errorMessage = response.ErrorMessage ?? "Failed to close session",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }

        return response.Result ?? JsonSerializer.Serialize(new
        {
            success = true,
            sessionId,
            saved = save
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Creates a new empty Excel file AND opens a session in one operation.
    /// Returns sessionId that must be used for all subsequent operations.
    /// Directory must exist - will not be created automatically.
    /// </summary>
    private static string CreateSessionAsync(string path, bool show, TimeSpan timeout)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("path is required for 'create' action", nameof(path));
        }

        // Validate Windows path format before any file operations
        var pathError = ExcelToolsBase.ValidateWindowsPath(path);
        if (pathError != null)
        {
            return pathError;
        }

        // Determine if macro-enabled from extension
        bool macroEnabled = path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase);
        var extension = macroEnabled ? ".xlsm" : ".xlsx";
        if (!path.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
        {
            path = Path.ChangeExtension(path, extension);
        }

        var timeoutSeconds = (int)timeout.TotalSeconds;
        var response = ExcelToolsBase.SendToServiceAsync(
            "session.create",
            null,
            new { filePath = path, macroEnabled, show = show, timeoutSeconds },
            timeoutSeconds
        ).GetAwaiter().GetResult();

        if (!response.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = response.ErrorMessage ?? "Failed to create session",
                filePath = path,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }

        return response.Result ?? JsonSerializer.Serialize(new
        {
            success = true,
            filePath = path,
            macroEnabled,
            show
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Closes the workbook (no-op with new single-instance architecture).
    /// LLM Pattern: This action is kept for backward compatibility but does nothing.
    /// With single-instance sessions, workbooks are automatically closed after each operation.
    /// </summary>
    private static string CloseWorkbook(string path)
    {
        return JsonSerializer.Serialize(new
        {
            success = true,
            filePath = path
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Lists all active sessions with status info. Lightweight operation - no Excel COM calls.
    /// LLM Pattern: Use this to verify sessions and check for running operations before closing.
    /// </summary>
    private static string ListSessions()
    {
        var response = ExcelToolsBase.SendToServiceAsync("session.list").GetAwaiter().GetResult();

        if (!response.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = response.ErrorMessage ?? "Failed to list sessions",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }

        return response.Result ?? JsonSerializer.Serialize(new
        {
            success = true,
            sessions = Array.Empty<object>(),
            count = 0
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Tests if an Excel file exists and is valid without opening it via Excel COM.
    /// LLM Pattern: Use this for discovery/connectivity testing before running operations.
    /// </summary>
    private static string TestFileAsync(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("path is required for 'test' action", nameof(path));
        }

        // Validate Windows path format before any file operations
        var pathError = ExcelToolsBase.ValidateWindowsPath(path);
        if (pathError != null)
        {
            return pathError;
        }

        var fileCommands = new FileCommands();
        var info = fileCommands.Test(path);

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






