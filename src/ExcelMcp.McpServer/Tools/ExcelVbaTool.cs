using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.ComInterop.Session;
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

RUN PARAMETERS:
- Format: 'Module.Procedure' (e.g., 'DataProcessor.ProcessData')
- Parameters: Comma-separated values passed to VBA procedure
- Example: moduleName='Module1.Calculate', parameters='Sheet1,A1:C10'

RELATED TOOLS:
- excel_file: Create .xlsm files for VBA automation")]
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
        string? parameters = null)
    {
        try
        {
            var vbaCommands = new VbaCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                VbaAction.List => await ListVbaScriptsAsync(vbaCommands, excelPath),
                VbaAction.View => await ViewVbaScriptAsync(vbaCommands, excelPath, moduleName),
                VbaAction.Export => await ExportVbaScriptAsync(vbaCommands, excelPath, moduleName, targetPath),
                VbaAction.Import => await ImportVbaScriptAsync(vbaCommands, excelPath, moduleName, sourcePath),
                VbaAction.Update => await UpdateVbaScriptAsync(vbaCommands, excelPath, moduleName, sourcePath),
                VbaAction.Run => await RunVbaScriptAsync(vbaCommands, excelPath, moduleName, parameters),
                VbaAction.Delete => await DeleteVbaScriptAsync(vbaCommands, excelPath, moduleName),
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

    private static async Task<string> ListVbaScriptsAsync(VbaCommands commands, string filePath)
    {
        var result = await commands.ListAsync(filePath);

        // Always return JSON (success or failure) - MCP clients handle the success flag
        var moduleCount = result.Scripts?.Count ?? 0;
        return JsonSerializer.Serialize(new
        {
            success = result.Success,
            errorMessage = result.ErrorMessage,
            scripts = result.Scripts,
            count = moduleCount,
            workflowHint = moduleCount == 0
                ? "No VBA modules found. Use 'import' to add VBA code."
                : $"Found {moduleCount} VBA module(s). Use 'view' to inspect or 'run' to execute."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ViewVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ModelContextProtocol.McpException("moduleName is required for view action");

        var result = await commands.ViewAsync(filePath, moduleName);

        var lineCount = result.Code?.Split('\n').Length ?? 0;
        var procedureCount = result.Procedures?.Count ?? 0;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            result.ModuleName,
            result.ModuleType,
            result.Code,
            result.LineCount,
            result.Procedures,
            workflowHint = $"Module '{moduleName}' has {lineCount} lines and {procedureCount} procedure(s)."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ExportVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName, string? targetPath)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(targetPath))
            throw new ModelContextProtocol.McpException("moduleName and targetPath are required for export action");

        var result = await commands.ExportAsync(filePath, moduleName, targetPath);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName,
            FilePath = targetPath,
            workflowHint = $"VBA module '{moduleName}' exported to {targetPath}."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ImportVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName, string? sourcePath)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ModelContextProtocol.McpException("moduleName is required for import action");
        if (string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("sourcePath is required for import action");

        var result = await commands.ImportAsync(filePath, moduleName, sourcePath);

        // Save the workbook after import
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName,
            SourcePath = sourcePath,
            workflowHint = $"VBA module '{moduleName}' imported from {sourcePath}. Ready to run."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName, string? sourcePath)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(sourcePath))
            throw new ModelContextProtocol.McpException("moduleName and sourcePath are required for update action");

        var result = await commands.UpdateAsync(filePath, moduleName, sourcePath);

        // Save the workbook after update
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName,
            SourcePath = sourcePath,
            workflowHint = $"VBA module '{moduleName}' updated from {sourcePath}. Changes saved."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RunVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName, string? parameters)
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

        var result = await commands.RunAsync(filePath, moduleName, null, paramArray);
        var paramCount = paramArray.Length;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ProcedureName = moduleName,
            ParameterCount = paramCount,
            workflowHint = $"VBA procedure '{moduleName}' executed with {paramCount} parameter(s)."
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteVbaScriptAsync(VbaCommands commands, string filePath, string? moduleName)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ModelContextProtocol.McpException("moduleName is required for delete action");

        var result = await commands.DeleteAsync(filePath, moduleName);

        // Save the workbook after delete
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName,
            workflowHint = $"VBA module '{moduleName}' deleted permanently. Changes saved."
        }, ExcelToolsBase.JsonOptions);
    }
}

