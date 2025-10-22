using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

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
[McpServerToolType]
public static class ExcelVbaTool
{
    /// <summary>
    /// Manage Excel VBA scripts - modules, procedures, and macro execution (requires .xlsm files)
    /// </summary>
    [McpServerTool(Name = "excel_vba")]
    [Description("Manage Excel VBA scripts and macros (requires .xlsm files). Supports: list, export, import, update, run, delete.")]
    public static string ExcelVba(
        [Required]
        [RegularExpression("^(list|export|import|update|run|delete)$")]
        [Description("Action: list, export, import, update, run, delete")]
        string action,

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
            var scriptCommands = new ScriptCommands();

            switch (action.ToLowerInvariant())
            {
                case "list":
                    return ListVbaScripts(scriptCommands, excelPath);
                case "export":
                    return ExportVbaScript(scriptCommands, excelPath, moduleName, targetPath);
                case "import":
                    return ImportVbaScript(scriptCommands, excelPath, moduleName, sourcePath);
                case "update":
                    return UpdateVbaScript(scriptCommands, excelPath, moduleName, sourcePath);
                case "run":
                    return RunVbaScript(scriptCommands, excelPath, moduleName, parameters);
                case "delete":
                    return DeleteVbaScript(scriptCommands, excelPath, moduleName);
                default:
                    ExcelToolsBase.ThrowUnknownAction(action, "list", "export", "import", "update", "run", "delete");
                    throw new InvalidOperationException(); // Never reached
            }
        }
        catch (ModelContextProtocol.McpException)
        {
            throw;
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action, excelPath);
            throw;
        }
    }

    private static string ListVbaScripts(ScriptCommands commands, string filePath)
    {
        var result = commands.List(filePath);

        // If listing failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Ensure VBA trust is enabled (run setup-vba-trust)",
                "Check that the file is .xlsm (macro-enabled)",
                "Verify the file exists and is accessible"
            };
            result.WorkflowHint = "List failed. Ensure VBA trust is enabled and file is .xlsm.";
            throw new ModelContextProtocol.McpException($"list failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'export' to backup VBA code",
            "Use 'run' to execute a VBA procedure",
            "Use 'import' to add new VBA modules"
        };
        result.WorkflowHint = "VBA modules listed. Next, export, run, or import as needed.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ExportVbaScript(ScriptCommands commands, string filePath, string? moduleName, string? vbaFilePath)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(vbaFilePath))
            throw new ModelContextProtocol.McpException("moduleName and vbaFilePath are required for export action");

        var result = commands.Export(filePath, moduleName, vbaFilePath).GetAwaiter().GetResult();

        // If export failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the module name exists",
                "Verify the target path is writable",
                "Use 'list' to see available modules"
            };
            result.WorkflowHint = "Export failed. Ensure the module exists and path is writable.";
            throw new ModelContextProtocol.McpException($"export failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Edit the exported VBA code as needed",
            "Use 'update' to re-import modified code",
            "Version control the exported .vba file"
        };
        result.WorkflowHint = "VBA module exported. Next, edit and update as needed.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ImportVbaScript(ScriptCommands commands, string filePath, string? moduleName, string? vbaFilePath)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(vbaFilePath))
            throw new ModelContextProtocol.McpException("moduleName and vbaFilePath are required for import action");

        var result = commands.Import(filePath, moduleName, vbaFilePath).GetAwaiter().GetResult();

        // If import failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the VBA file exists",
                "Verify VBA trust is enabled",
                "Ensure the module name doesn't already exist"
            };
            result.WorkflowHint = "Import failed. Ensure the file exists and module is unique.";
            throw new ModelContextProtocol.McpException($"import failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'list' to verify the import",
            "Use 'run' to execute procedures in the module",
            "Test the imported VBA code"
        };
        result.WorkflowHint = "VBA module imported. Next, verify and test.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string UpdateVbaScript(ScriptCommands commands, string filePath, string? moduleName, string? vbaFilePath)
    {
        if (string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(vbaFilePath))
            throw new ModelContextProtocol.McpException("moduleName and vbaFilePath are required for update action");

        var result = commands.Update(filePath, moduleName, vbaFilePath).GetAwaiter().GetResult();

        // If update failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the module exists",
                "Verify the VBA file exists and is accessible",
                "Use 'list' to see available modules"
            };
            result.WorkflowHint = "Update failed. Ensure the module and file exist.";
            throw new ModelContextProtocol.McpException($"update failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'run' to test the updated code",
            "Use 'export' to backup the updated module",
            "Verify the code changes work as expected"
        };
        result.WorkflowHint = "VBA module updated. Next, test and verify.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string RunVbaScript(ScriptCommands commands, string filePath, string? moduleName, string? parameters)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ModelContextProtocol.McpException("moduleName (format: 'Module.Procedure') is required for run action");

        // Parse parameters if provided
        var paramArray = string.IsNullOrEmpty(parameters)
            ? Array.Empty<string>()
            : parameters.Split(',', StringSplitOptions.RemoveEmptyEntries)
                       .Select(p => p.Trim())
                       .ToArray();

        var result = commands.Run(filePath, moduleName, paramArray);

        // If VBA execution failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the procedure exists (format: 'Module.Procedure')",
                "Verify the parameters are correct",
                "Review the VBA code for errors",
                "Ensure VBA trust is enabled"
            };
            result.WorkflowHint = "VBA run failed. Ensure the procedure exists and parameters are correct.";
            throw new ModelContextProtocol.McpException($"run failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use worksheet 'read' to verify VBA made expected changes",
            "Review output or return values",
            "Run again with different parameters if needed"
        };
        result.WorkflowHint = "VBA executed successfully. Next, verify results.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteVbaScript(ScriptCommands commands, string filePath, string? moduleName)
    {
        if (string.IsNullOrEmpty(moduleName))
            throw new ModelContextProtocol.McpException("moduleName is required for delete action");

        var result = commands.Delete(filePath, moduleName);

        // If delete failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the module exists",
                "Use 'list' to see available modules",
                "Verify the module name is correct"
            };
            result.WorkflowHint = "Delete failed. Ensure the module exists and name is correct.";
            throw new ModelContextProtocol.McpException($"delete failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'list' to verify the deletion",
            "Export other modules for backup",
            "Review remaining VBA code"
        };
        result.WorkflowHint = "VBA module deleted. Next, verify and backup remaining code.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
