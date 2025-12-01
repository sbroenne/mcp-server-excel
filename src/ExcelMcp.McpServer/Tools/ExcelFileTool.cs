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
    /// Session lifecycle: open (get sessionId) then use sessionId with other tools then close (save:true to persist, save:false to discard).
    /// No separate 'save' action exists - use close with save:true to persist changes.
    /// Keep session open across ALL operations - only close when workflow complete and all actions have reported a status.
    /// Supports .xlsx (standard) and .xlsm (macro-enabled) formats.
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="excelPath">Excel file path (.xlsx or .xlsm) - required for open/create-empty, not used for close</param>
    /// <param name="sessionId">Session ID from 'open' action - required for close</param>
    /// <param name="save">Save changes before closing (default: false)</param>
    [McpServerTool(Name = "excel_file")]
    [McpMeta("category", "session")]
    public static partial string ExcelFile(
        FileAction action,
        string? excelPath,
        string? sessionId,
        bool save)
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
                    FileAction.Open => OpenSessionAsync(excelPath!),
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
    private static string OpenSessionAsync(string excelPath)
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
            string sessionId = ExcelToolsBase.GetSessionManager().CreateSession(excelPath);

            return JsonSerializer.Serialize(new
            {
                success = true,
                sessionId,
                filePath = excelPath
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

        try
        {
            bool success = ExcelToolsBase.GetSessionManager().CloseSession(sessionId, save);

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

