using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel file management tool for MCP server.
/// Manages Excel file creation and session lifecycle for automation workflows.
/// Supports .xlsx (standard) and .xlsm (macro-enabled) formats.
/// </summary>
[McpServerToolType]
[SuppressMessage("Performance", "CA1861:Avoid constant arrays as arguments", Justification = "Simple workflow arrays in sealed static class")]
public static class ExcelFileTool
{
    /// <summary>
    /// Create new Excel files for automation workflows
    /// </summary>
    [McpServerTool(Name = "excel_file")]
    [Description(@"Manage Excel files and sessions.

⚠️ SESSION LIFECYCLE REQUIRED:
1. OPEN - Start session, get sessionId
2. OPERATE - Use sessionId with other tools
3. SAVE - Explicitly save changes (optional)
4. CLOSE - End session, release resources

WORKFLOWS:
- Default: open → operations(sessionId) → save → close
- Read-only: open → read(sessionId) → close (no save)
- Discard changes: open → modify(sessionId) → close (no save)

FILE FORMATS:
- .xlsx:
- .xlsm:
")]
    public static async Task<string> ExcelFile(
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        FileAction action,

        [Description("Excel file path (.xlsx or .xlsm extension) - required for open/create-empty, not used for save/close")]
        string? excelPath = null,

        [Description("Session ID from 'open' action - required for save/close, not used for open/create-empty/test")]
        string? sessionId = null)
    {
        try
        {
            var fileCommands = new FileCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                FileAction.Open => await OpenSessionAsync(excelPath!),
                FileAction.Save => await SaveSessionAsync(sessionId!),
                FileAction.Close => await CloseSessionAsync(sessionId!),
                FileAction.CreateEmpty => await CreateEmptyFileAsync(fileCommands, excelPath!,
                    excelPath!.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)),
                FileAction.CloseWorkbook => CloseWorkbook(excelPath!),
                FileAction.Test => await TestFileAsync(fileCommands, excelPath!),
                _ => throw new ModelContextProtocol.McpException($"Unknown action: {action} ({action.ToActionString()})")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);
            throw; // Unreachable but satisfies compiler
        }
    }

    /// <summary>
    /// Opens an Excel file and creates a new session.
    /// Returns sessionId that must be used for all subsequent operations.
    /// </summary>
    private static async Task<string> OpenSessionAsync(string excelPath)
    {
        if (string.IsNullOrWhiteSpace(excelPath))
        {
            throw new ModelContextProtocol.McpException("excelPath is required for 'open' action");
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

        string sessionId = await ExcelToolsBase.GetSessionManager().CreateSessionAsync(excelPath);

        return JsonSerializer.Serialize(new
        {
            success = true,
            sessionId,
            filePath = excelPath,
            message = "Session opened successfully",
            suggestedNextActions = new List<string>
            {
                "Use the returned sessionId with other excel_* tools",
                "Call excel_file 'save' action to persist changes when ready",
                "Call excel_file 'close' action when finished to release resources"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Saves changes for an active session.
    /// Does not close the session - call 'close' action separately.
    /// </summary>
    private static async Task<string> SaveSessionAsync(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            throw new ModelContextProtocol.McpException("sessionId is required for 'save' action");
        }

        bool success = await ExcelToolsBase.GetSessionManager().SaveSessionAsync(sessionId);

        if (success)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                sessionId,
                message = "Changes saved successfully",
                suggestedNextActions = new List<string>
                {
                    "Call excel_file 'close' action when finished"
                }
            }, ExcelToolsBase.JsonOptions);
        }

        return JsonSerializer.Serialize(new
        {
            success = false,
            sessionId,
            errorMessage = $"Session '{sessionId}' not found",
            isError = true,
            suggestedNextActions = new List<string>
            {
                "Session may have already been closed",
                "Use excel_file 'open' to start a new session"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Closes an active session without saving changes.
    /// To save before closing, call 'save' action first.
    /// </summary>
    private static async Task<string> CloseSessionAsync(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            throw new ModelContextProtocol.McpException("sessionId is required for 'close' action");
        }

        bool success = await ExcelToolsBase.GetSessionManager().CloseSessionAsync(sessionId);

        if (success)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                sessionId,
                message = "Session closed successfully (changes not saved)",
                suggestedNextActions = new List<string>
                {
                    "If changes needed to be saved, call 'save' before 'close' next time"
                }
            }, ExcelToolsBase.JsonOptions);
        }

        return JsonSerializer.Serialize(new
        {
            success = false,
            sessionId,
            errorMessage = $"Session '{sessionId}' not found",
            isError = true,
            message = "Session not found or already closed"
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Creates a new empty Excel file (.xlsx or .xlsm based on macroEnabled flag).
    /// LLM Pattern: Use this when you need a fresh Excel workbook for automation.
    /// </summary>
    private static async Task<string> CreateEmptyFileAsync(FileCommands fileCommands, string excelPath, bool macroEnabled)
    {
        var extension = macroEnabled ? ".xlsm" : ".xlsx";
        if (!excelPath.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
        {
            excelPath = Path.ChangeExtension(excelPath, extension);
        }

        var result = await fileCommands.CreateEmptyAsync(excelPath, overwriteIfExists: false);

        if (result.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                filePath = result.FilePath,
                macroEnabled,
                message = "Excel file created successfully",
                suggestedNextActions = new List<string>
                {
                    "Call excel_file 'open' to start a new session for this file",
                    "Then use other excel_* tools with the sessionId"
                }
            }, ExcelToolsBase.JsonOptions);
        }

        return JsonSerializer.Serialize(new
        {
            success = false,
            errorMessage = result.ErrorMessage,
            filePath = excelPath,
            message = result.ErrorMessage
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
            filePath = excelPath,
            message = "Workbook closure is automatic with single-instance architecture",
            suggestedNextActions = new List<string>
            {
                "Use other excel_* tools to work with files",
                "Each operation automatically manages its own Excel instance"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Tests if an Excel file exists and is valid without opening it via Excel COM.
    /// LLM Pattern: Use this for discovery/connectivity testing before running operations.
    /// </summary>
    private static async Task<string> TestFileAsync(FileCommands fileCommands, string excelPath)
    {
        if (string.IsNullOrWhiteSpace(excelPath))
        {
            throw new ModelContextProtocol.McpException("excelPath is required for 'test' action");
        }

        var result = await fileCommands.TestAsync(excelPath);

        if (result.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                filePath = result.FilePath,
                exists = result.Exists,
                isValid = result.IsValid,
                extension = result.Extension,
                size = result.Size,
                lastModified = result.LastModified,
                message = "File exists and is a valid Excel file",
                suggestedNextActions = new List<string>
                {
                    "Use excel_powerquery to manage Power Query connections",
                    "Use excel_vba to manage VBA macros",
                    "Use begin_excel_batch for multi-operation workflows"
                }
            }, ExcelToolsBase.JsonOptions);
        }

        return JsonSerializer.Serialize(new
        {
            success = false,
            filePath = result.FilePath,
            exists = result.Exists,
            isValid = result.IsValid,
            extension = result.Extension,
            size = result.Size,
            lastModified = result.LastModified,
            errorMessage = result.ErrorMessage,
            message = result.ErrorMessage,
            isError = true
        }, ExcelToolsBase.JsonOptions);
    }
}
