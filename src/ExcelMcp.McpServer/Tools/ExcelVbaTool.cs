using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel VBA script management tool for MCP server.
/// Manages VBA macro operations, code import/export, and script execution in macro-enabled workbooks.
///
/// IMPORTANT: Requires .xlsm files! VBA operations only work with macro-enabled Excel files.
///
/// Prerequisites: VBA trust must be enabled for automation. Use setup-vba-trust command to configure.
/// </summary>
[McpServerToolType]
public static partial class ExcelVbaTool
{
    /// <summary>
    /// Manage Excel VBA scripts and macros (requires .xlsm files).
    /// REQUIREMENTS: File format must be .xlsm (macro-enabled) only. VBA trust must be enabled in Excel settings (one-time setup).
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="excelPath">Excel file path (must be .xlsm for VBA operations)</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action (required for all VBA operations)</param>
    /// <param name="moduleName">VBA module name or procedure name (format: 'Module.Procedure' for run)</param>
    /// <param name="vbaCode">VBA code content as string (for import/update actions)</param>
    /// <param name="parameters">Parameters for VBA procedure execution (comma-separated)</param>
    [McpServerTool(Name = "excel_vba")]
    [McpMeta("category", "automation")]
    public static partial string ExcelVba(
        VbaAction action,
        string excelPath,
        string sessionId,
        [DefaultValue(null)] string? moduleName,
        [DefaultValue(null)] string? vbaCode,
        [DefaultValue(null)] string? parameters)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_vba",
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

        ExcelToolsBase.WithSession(
            sessionId,
            batch =>
            {
                commands.Import(batch, moduleName, vbaCode);
                return 0;
            });

        return JsonSerializer.Serialize(new
        {
            success = true,
            moduleName,
            message = $"Imported VBA module '{moduleName}'."
        }, ExcelToolsBase.JsonOptions);
    }

    private static string UpdateVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName, string? vbaCode)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(vbaCode))
            throw new ArgumentException("moduleName and vbaCode are required for update action", "moduleName,vbaCode");

        ExcelToolsBase.WithSession(
            sessionId,
            batch =>
            {
                commands.Update(batch, moduleName, vbaCode);
                return 0;
            });

        return JsonSerializer.Serialize(new
        {
            success = true,
            moduleName,
            message = $"Updated VBA module '{moduleName}'."
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

        ExcelToolsBase.WithSession(
            sessionId,
            batch =>
            {
                commands.Run(batch, moduleName, null, paramArray);
                return 0;
            });

        var paramCount = paramArray.Length;

        return JsonSerializer.Serialize(new
        {
            success = true,
            procedureName = moduleName,
            parameterCount = paramCount,
            message = $"Executed VBA procedure '{moduleName}'."
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteVbaScriptAsync(VbaCommands commands, string sessionId, string? moduleName)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ArgumentException("moduleName is required for delete action", nameof(moduleName));

        ExcelToolsBase.WithSession(
            sessionId,
            batch =>
            {
                commands.Delete(batch, moduleName);
                return 0;
            });

        return JsonSerializer.Serialize(new
        {
            success = true,
            moduleName,
            message = $"Deleted VBA module '{moduleName}'."
        }, ExcelToolsBase.JsonOptions);
    }
}

