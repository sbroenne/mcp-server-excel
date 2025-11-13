using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - workflow hints are contextual per-call

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel VBA script management tool for MCP server.
/// Manages VBA macro operations, code import/export, and script execution in macro-enabled workbooks.
///
/// ⚠️ IMPORTANT: Requires .xlsm files! VBA operations only work with macro-enabled Excel files.
///
/// Prerequisites: VBA trust must be enabled for automation. Use setup-vba-trust command to configure.
/// </summary>
[McpServerToolType]
[SuppressMessage("Performance", "CA1861:Avoid constant arrays as arguments", Justification = "Simple workflow arrays in sealed static class")]
public static class ExcelVbaTool
{
    /// <summary>
    /// Manage Excel VBA scripts - modules, procedures, and macro execution (requires .xlsm files)
    /// </summary>
    [McpServerTool(Name = "excel_vba")]
    [Description(@"Manage Excel VBA scripts and macros (requires .xlsm files).

⚠️ REQUIREMENTS:
- File format: .xlsm (macro-enabled) only
- VBA trust: Must be enabled in Excel settings (one-time setup)
")]
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
                : ["Use 'run' to execute macros", "Use 'export' to backup VBA code", "Use 'view' to inspect module code"]
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

        var inBatch = !string.IsNullOrEmpty(batchId);
        var lineCount = result.Code?.Split('\n').Length ?? 0;
        var procedureCount = result.Procedures?.Count ?? 0;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ModuleName,
            result.ModuleType,
            result.Code,
            result.LineCount,
            result.Procedures,
            workflowHint = $"Module '{moduleName}' has {lineCount} lines and {procedureCount} procedure(s).",
            suggestedNextActions = new[]
            {
                "Use 'run' to execute procedures from this module",
                "Use 'export' to save VBA code to file for version control",
                "Use 'update' to modify the module code",
                inBatch ? "View more modules in this batch" : "Viewing multiple modules? Use excel_batch for efficiency"
            }
        }, ExcelToolsBase.JsonOptions);
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

        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName,
            FilePath = targetPath,
            workflowHint = $"VBA module '{moduleName}' exported to {targetPath}.",
            suggestedNextActions = new[]
            {
                "Commit exported VBA file to version control",
                "Use 'import' or 'update' to restore VBA code from backup",
                "Review exported code for documentation or code review",
                inBatch ? "Export more modules in this batch" : "Exporting multiple modules? Use excel_batch for efficiency"
            }
        }, ExcelToolsBase.JsonOptions);
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

        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName,
            SourcePath = sourcePath,
            workflowHint = $"VBA module '{moduleName}' imported from {sourcePath}. Ready to run.",
            suggestedNextActions = new[]
            {
                "Use 'view' to inspect the imported VBA code",
                "Use 'run' to execute procedures from this module",
                "Use 'list' to see all VBA modules including the new one",
                inBatch ? "Import more modules in this batch" : "Importing multiple modules? Use excel_batch for efficiency"
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

        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName,
            SourcePath = sourcePath,
            workflowHint = $"VBA module '{moduleName}' updated from {sourcePath}. Changes saved.",
            suggestedNextActions = new[]
            {
                "Use 'view' to verify the updated VBA code",
                "Use 'run' to test the updated procedures",
                "Use 'export' to backup the updated module",
                inBatch ? "Update more modules in this batch" : "Updating multiple modules? Use excel_batch for efficiency"
            }
        }, ExcelToolsBase.JsonOptions);
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
            async (batch) => await commands.RunAsync(batch, moduleName, null, paramArray));

        var inBatch = !string.IsNullOrEmpty(batchId);
        var paramCount = paramArray.Length;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ProcedureName = moduleName,
            ParameterCount = paramCount,
            workflowHint = $"VBA procedure '{moduleName}' executed with {paramCount} parameter(s).",
            suggestedNextActions = new[]
            {
                "Check Excel workbook for procedure output (worksheets, cells, etc.)",
                "Use excel_range or excel_worksheet to verify VBA changes",
                "Use 'view' to inspect VBA code if unexpected results",
                inBatch ? "Run more procedures in this batch" : "Running multiple procedures? Use excel_batch for efficiency"
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

        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName,
            workflowHint = $"VBA module '{moduleName}' deleted permanently. Changes saved.",
            suggestedNextActions = new[]
            {
                "Use 'list' to verify module was removed",
                "Use 'import' to restore module from backup if needed",
                "Export remaining modules for backup before further deletions",
                inBatch ? "Delete more modules in this batch" : "Deleting multiple modules? Use excel_batch for efficiency"
            }
        }, ExcelToolsBase.JsonOptions);
    }
}
