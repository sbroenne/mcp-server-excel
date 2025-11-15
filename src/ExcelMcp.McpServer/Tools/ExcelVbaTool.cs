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

        [Required]
        [Description("Session ID from excel_file 'open' action (required for all VBA operations)")]
        string sessionId,

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
                VbaAction.List => await ListVbaScriptsAsync(vbaCommands, sessionId),
                VbaAction.View => await ViewVbaScriptAsync(vbaCommands, sessionId, moduleName),
                VbaAction.Export => await ExportVbaScriptAsync(vbaCommands, sessionId, moduleName, targetPath),
                VbaAction.Import => await ImportVbaScriptAsync(vbaCommands, sessionId, moduleName, sourcePath),
                VbaAction.Update => await UpdateVbaScriptAsync(vbaCommands, sessionId, moduleName, sourcePath),
                VbaAction.Run => await RunVbaScriptAsync(vbaCommands, sessionId, moduleName, parameters),
                VbaAction.Delete => await DeleteVbaScriptAsync(vbaCommands, sessionId, moduleName),
                _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"{action.ToActionString()} failed: {ex.Message}",
                filePath = excelPath,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static async Task<string> ListVbaScriptsAsync(VbaCommands commands, string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            commands.ListAsync);

        // If listing failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        var moduleCount = result.Scripts?.Count ?? 0;
        return JsonSerializer.Serialize(new
        {
            success = true,
            scripts = result.Scripts,
            count = moduleCount
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ViewVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ArgumentException("moduleName is required for view action", nameof(moduleName));

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ViewAsync(batch, moduleName));

        var lineCount = result.Code?.Split('\n').Length ?? 0;
        var procedureCount = result.Procedures?.Count ?? 0;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ModuleName,
            result.ModuleType,
            result.Code,
            result.LineCount,
            result.Procedures
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ExportVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName, string? targetPath)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ArgumentException("moduleName is required for export action", nameof(moduleName));

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ViewAsync(batch, moduleName));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ImportVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName, string? sourcePath)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ArgumentException("moduleName is required for import action", nameof(moduleName));
        if (string.IsNullOrEmpty(sourcePath))
            throw new ArgumentException("sourcePath is required for import action", nameof(sourcePath));

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ImportAsync(batch, moduleName, sourcePath));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName,
            SourcePath = sourcePath
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName, string? sourcePath)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(sourcePath))
            throw new ArgumentException("moduleName and sourcePath are required for update action", "moduleName,sourcePath");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.UpdateAsync(batch, moduleName, sourcePath));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName,
            SourcePath = sourcePath
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RunVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName, string? parameters)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ArgumentException("moduleName (format: 'Module.Procedure') is required for run action", nameof(moduleName));

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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.RunAsync(batch, moduleName, null, paramArray));
        var paramCount = paramArray.Length;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ProcedureName = moduleName,
            ParameterCount = paramCount
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ArgumentException("moduleName is required for delete action", nameof(moduleName));

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.DeleteAsync(batch, moduleName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName
        }, ExcelToolsBase.JsonOptions);
    }
}
