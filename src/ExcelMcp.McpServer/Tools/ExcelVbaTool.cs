using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel VBA script management tool for MCP server.
/// Handles VBA macro operations, code management, and script execution.
///
/// ⚠️ IMPORTANT: Requires .xlsm files! VBA operations only work with macro-enabled Excel files.
///
/// LLM Usage Patterns:
/// - Use "list" to see all VBA modules and procedures
/// - Use "view" to inspect VBA code without exporting
/// - Use "export" to backup VBA code to .vba files
/// - Use "import" to load VBA modules from files
/// - Use "update" to modify existing VBA modules
/// - Use "run" to execute VBA macros with parameters
/// - Use "delete" to remove VBA modules
///
/// Setup Required: Run setup-vba-trust command once before using VBA operations.
/// </summary>
[McpServerToolType]
public static class ExcelVbaTool
{
    /// <summary>
    /// Manage Excel VBA scripts - modules, procedures, and macro execution (requires .xlsm files)
    /// </summary>
    [McpServerTool(Name = "excel_vba")]
    [Description("Manage Excel VBA scripts and macros (requires .xlsm files). Supports: list, view, export, import, update, run, delete. Optional batchId for batch sessions.")]
    public static async Task<string> ExcelVba(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        VbaAction action,

        [Required]
        [FileExtensions(Extensions = "xlsm")]
        [Description("Excel file path (must be .xlsm for VBA operations)")]
        string excelPath,

        [StringLength(255, MinimumLength = 1)]
        [Description("VBA module name or procedure name (format: 'Module.Procedure' for run)")]
        string? moduleName = null,

        [FileExtensions(Extensions = "vba,bas,txt")]
        [Description("Source VBA file path (for import/update) or target file path (for export)")]
        string? sourcePath = null,

        [FileExtensions(Extensions = "vba,bas,txt")]
        [Description("Target VBA file path (for export action)")]
        string? targetPath = null,

        [Description("Parameters for VBA procedure execution (comma-separated)")]
        string? parameters = null,

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var vbaCommands = new VbaCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                VbaAction.List => await ListVbaScriptsAsync(vbaCommands, excelPath, batchId),
                VbaAction.View => await ViewVbaScriptAsync(vbaCommands, excelPath, moduleName, batchId),
                VbaAction.Export => await ExportVbaScriptAsync(vbaCommands, excelPath, moduleName, targetPath, batchId),
                VbaAction.Import => await ImportVbaScriptAsync(vbaCommands, excelPath, moduleName, sourcePath, batchId),
                VbaAction.Update => await UpdateVbaScriptAsync(vbaCommands, excelPath, moduleName, sourcePath, batchId),
                VbaAction.Run => await RunVbaScriptAsync(vbaCommands, excelPath, moduleName, parameters, batchId),
                VbaAction.Delete => await DeleteVbaScriptAsync(vbaCommands, excelPath, moduleName, batchId),
                _ => throw new ModelContextProtocol.McpException($"Unknown action: {action} ({action.ToActionString()})")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw;
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);
            throw;
        }
    }

    private static async Task<string> ListVbaScriptsAsync(VbaCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListAsync(batch));

        // If listing failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        var moduleCount = result.Scripts?.Count ?? 0;
        return JsonSerializer.Serialize(new
        {
            success = true,
            scripts = result.Scripts,
            count = moduleCount,
            workflowHint = moduleCount == 0
                ? "No VBA modules found. Use 'import' to add VBA code."
                : $"Found {moduleCount} VBA module(s). Use 'view' to inspect or 'run' to execute.",
            suggestedNextActions = moduleCount == 0
                ? new[] { "Use 'import' to add VBA modules from .vba files", "Use excel_file to create .xlsm files for VBA" }
                : new[] { "Use 'run' to execute macros", "Use 'export' to backup VBA code", "Use 'view' to inspect module code" }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ViewVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName, string? batchId)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ModelContextProtocol.McpException("moduleName is required for view action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ViewAsync(batch, moduleName));

        // If view failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ExportVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName, string? targetPath, string? batchId)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(targetPath))
            throw new ModelContextProtocol.McpException("moduleName and targetPath are required for export action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ExportAsync(batch, moduleName, targetPath));

        // If export failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ImportVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName, string? sourcePath, string? batchId)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ModelContextProtocol.McpException("moduleName is required for import action");
        if (string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("sourcePath is required for import action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.ImportAsync(batch, moduleName, sourcePath));

        // If import failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            success = true,
            moduleName,
            sourcePath,
            message = $"VBA module '{moduleName}' imported successfully",
            workflowHint = "Module imported. Use 'run' to execute or 'view' to inspect code.",
            suggestedNextActions = new[]
            {
                batchId != null
                    ? $"Continue using batchId '{batchId}' to import more modules"
                    : "Use excel_batch for importing multiple modules (75-90% faster)",
                "Use 'run' action to execute the imported macro",
                "Use 'view' to verify the imported code"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName, string? sourcePath, string? batchId)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("moduleName and sourcePath are required for update action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.UpdateAsync(batch, moduleName, sourcePath));

        // If update failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RunVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName, string? parameters, string? batchId)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ModelContextProtocol.McpException("moduleName (format: 'Module.Procedure') is required for run action");

        // Parse parameters if provided
        string[] paramArray;
        if (string.IsNullOrEmpty(parameters))
        {
            paramArray = Array.Empty<string>();
        }
        else
        {
            paramArray = parameters.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                   .Select(p => p.Trim())
                                   .ToArray();
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false, // VBA execution doesn't save unless VBA code does
            async (batch) => await commands.RunAsync(batch, moduleName, paramArray));

        // If VBA execution failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            success = true,
            moduleName,
            parameters = paramArray,
            message = "VBA procedure executed successfully",
            workflowHint = "Macro executed. Check results with excel_range if data was modified.",
            suggestedNextActions = new[]
            {
                "Use excel_range 'get-values' to verify data changes",
                "Use 'list' to see all available macros",
                "Save workbook if macro made changes"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName, string? batchId)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ModelContextProtocol.McpException("moduleName is required for delete action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.DeleteAsync(batch, moduleName));

        // If delete failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
