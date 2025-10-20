using Sbroenne.ExcelMcp.Core.Commands;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel VBA script management tool for MCP server.
/// Handles VBA macro operations, code management, and script execution.
/// 
/// ⚠️ IMPORTANT: Requires .xlsm files! VBA operations only work with macro-enabled Excel files.
/// 
/// LLM Usage Patterns:
/// - Use "list" to see all VBA modules and procedures
/// - Use "export" to backup VBA code to .vba files  
/// - Use "import" to load VBA modules from files
/// - Use "update" to modify existing VBA modules
/// - Use "run" to execute VBA macros with parameters
/// - Use "delete" to remove VBA modules
/// 
/// Setup Required: Run setup-vba-trust command once before using VBA operations.
/// </summary>
public static class ExcelVbaTool
{
    /// <summary>
    /// Manage Excel VBA scripts - modules, procedures, and macro execution (requires .xlsm files)
    /// </summary>
    [McpServerTool(Name = "excel_vba")]
    [Description("Manage Excel VBA scripts and macros (requires .xlsm files). Supports: list, export, import, update, run, delete.")]
    public static string ExcelVba(
        [Description("Action: list, export, import, update, run, delete")] string action,
        [Description("Excel file path (must be .xlsm for VBA operations)")] string filePath,
        [Description("VBA module name or procedure name (format: 'Module.Procedure' for run)")] string? moduleName = null,
        [Description("VBA file path (.vba extension for import/export/update)")] string? vbaFilePath = null,
        [Description("Parameters for VBA procedure execution (comma-separated)")] string? parameters = null)
    {
        try
        {
            var scriptCommands = new ScriptCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ListVbaScripts(scriptCommands, filePath),
                "export" => ExportVbaScript(scriptCommands, filePath, moduleName, vbaFilePath),
                "import" => ImportVbaScript(scriptCommands, filePath, moduleName, vbaFilePath),
                "update" => UpdateVbaScript(scriptCommands, filePath, moduleName, vbaFilePath),
                "run" => RunVbaScript(scriptCommands, filePath, moduleName, parameters),
                "delete" => DeleteVbaScript(scriptCommands, filePath, moduleName),
                _ => ExcelToolsBase.CreateUnknownActionError(action, "list", "export", "import", "update", "run", "delete")
            };
        }
        catch (Exception ex)
        {
            return ExcelToolsBase.CreateExceptionError(ex, action, filePath);
        }
    }

    private static string ListVbaScripts(ScriptCommands commands, string filePath)
    {
        var result = commands.List(filePath);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ExportVbaScript(ScriptCommands commands, string filePath, string? moduleName, string? vbaFilePath)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(vbaFilePath))
            return JsonSerializer.Serialize(new { error = "moduleName and vbaFilePath are required for export action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Export(filePath, moduleName, vbaFilePath).GetAwaiter().GetResult();
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ImportVbaScript(ScriptCommands commands, string filePath, string? moduleName, string? vbaFilePath)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(vbaFilePath))
            return JsonSerializer.Serialize(new { error = "moduleName and vbaFilePath are required for import action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Import(filePath, moduleName, vbaFilePath).GetAwaiter().GetResult();
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string UpdateVbaScript(ScriptCommands commands, string filePath, string? moduleName, string? vbaFilePath)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(vbaFilePath))
            return JsonSerializer.Serialize(new { error = "moduleName and vbaFilePath are required for update action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Update(filePath, moduleName, vbaFilePath).GetAwaiter().GetResult();
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string RunVbaScript(ScriptCommands commands, string filePath, string? moduleName, string? parameters)
    {
        if (string.IsNullOrEmpty(moduleName))
            return JsonSerializer.Serialize(new { error = "moduleName (format: 'Module.Procedure') is required for run action" }, ExcelToolsBase.JsonOptions);

        // Parse parameters if provided
        var paramArray = string.IsNullOrEmpty(parameters) 
            ? Array.Empty<string>() 
            : parameters.Split(',', StringSplitOptions.RemoveEmptyEntries)
                       .Select(p => p.Trim())
                       .ToArray();

        var result = commands.Run(filePath, moduleName, paramArray);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteVbaScript(ScriptCommands commands, string filePath, string? moduleName)
    {
        if (string.IsNullOrEmpty(moduleName))
            return JsonSerializer.Serialize(new { error = "moduleName is required for delete action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Delete(filePath, moduleName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}