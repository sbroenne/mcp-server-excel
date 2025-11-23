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
    public static string ExcelVba(
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

        [Description("VBA code content as string (for import/update actions)")]
        string? vbaCode = null,

        [Description("Parameters for VBA procedure execution (comma-separated)")]
        string? parameters = null)
    {
        return ExcelToolsBase.ExecuteToolAction(
            action.ToActionString(),
            excelPath,
            () =>
            {
                var vbaCommands = new VbaCommands();

                // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
                return action switch
                {
                    VbaAction.List => ListVbaScriptsAsync(vbaCommands, sessionId),
                    VbaAction.View => ViewVbaScriptAsync(vbaCommands, sessionId, moduleName),
                    VbaAction.Import => ImportVbaScriptAsync(vbaCommands, sessionId, moduleName, vbaCode),
                    VbaAction.Update => UpdateVbaScriptAsync(vbaCommands, sessionId, moduleName, vbaCode),
                    VbaAction.Run => RunVbaScriptAsync(vbaCommands, sessionId, moduleName, parameters),
                    VbaAction.Delete => DeleteVbaScriptAsync(vbaCommands, sessionId, moduleName),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListVbaScriptsAsync(VbaCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.List(batch));

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

    private static string ViewVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ArgumentException("moduleName is required for view action", nameof(moduleName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.View(batch, moduleName));

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

    private static string ImportVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName, string? vbaCode)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ArgumentException("moduleName is required for import action", nameof(moduleName));
        if (string.IsNullOrEmpty(vbaCode))
            throw new ArgumentException("vbaCode is required for import action", nameof(vbaCode));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Import(batch, moduleName, vbaCode));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName
        }, ExcelToolsBase.JsonOptions);
    }

    private static string UpdateVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName, string? vbaCode)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(vbaCode))
            throw new ArgumentException("moduleName and vbaCode are required for update action", "moduleName,vbaCode");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Update(batch, moduleName, vbaCode));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RunVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName, string? parameters)
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

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Run(batch, moduleName, null, paramArray));
        var paramCount = paramArray.Length;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ProcedureName = moduleName,
            ParameterCount = paramCount
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ArgumentException("moduleName is required for delete action", nameof(moduleName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Delete(batch, moduleName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            ModuleName = moduleName
        }, ExcelToolsBase.JsonOptions);
    }
}

