using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel worksheet management tool for MCP server.
/// Handles worksheet operations, data reading/writing, and sheet management.
///
/// LLM Usage Patterns:
/// - Use "list" to see all worksheets in a workbook
/// - Use "read" to extract data from worksheet ranges
/// - Use "write" to populate worksheets from CSV files
/// - Use "create" to add new worksheets
/// - Use "rename" to change worksheet names
/// - Use "copy" to duplicate worksheets
/// - Use "delete" to remove worksheets
/// - Use "clear" to empty worksheet ranges
/// - Use "append" to add data to existing worksheet content
/// </summary>
[McpServerToolType]
public static class ExcelWorksheetTool
{
    /// <summary>
    /// Manage Excel worksheets - data operations, sheet management, and content manipulation
    /// </summary>
    [McpServerTool(Name = "worksheet")]
    [Description("Manage Excel worksheets and data. Supports: list, read, write, create, rename, copy, delete, clear, append.")]
    public static string Worksheet(
        [Required]
        [RegularExpression("^(list|read|write|create|rename|copy|delete|clear|append)$")]
        [Description("Action: list, read, write, create, rename, copy, delete, clear, append")]
        string action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Worksheet name (required for most actions)")]
        string? sheetName = null,

        [Description("Excel range (e.g., 'A1:D10' for read/clear) or CSV file path (for write/append)")]
        string? range = null,

        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("New sheet name (for rename) or source sheet name (for copy)")]
        string? targetName = null)
    {
        try
        {
            var sheetCommands = new SheetCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ListWorksheets(sheetCommands, excelPath),
                "read" => ReadWorksheet(sheetCommands, excelPath, sheetName, range),
                "write" => WriteWorksheet(sheetCommands, excelPath, sheetName, range),
                "create" => CreateWorksheet(sheetCommands, excelPath, sheetName),
                "rename" => RenameWorksheet(sheetCommands, excelPath, sheetName, targetName),
                "copy" => CopyWorksheet(sheetCommands, excelPath, sheetName, targetName),
                "delete" => DeleteWorksheet(sheetCommands, excelPath, sheetName),
                "clear" => ClearWorksheet(sheetCommands, excelPath, sheetName, range),
                "append" => AppendWorksheet(sheetCommands, excelPath, sheetName, range),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{action}'. Supported: list, read, write, create, rename, copy, delete, clear, append")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action, excelPath);
            throw; // Unreachable but satisfies compiler
        }
    }

    private static string ListWorksheets(SheetCommands commands, string filePath)
    {
        var result = commands.List(filePath);

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the Excel file exists and is accessible",
                "Verify the file path is correct"
            };
            result.WorkflowHint = "List failed. Ensure the file exists and retry.";
            throw new ModelContextProtocol.McpException($"list failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'read' to extract data from a worksheet",
            "Use 'create' to add a new worksheet",
            "Use 'delete' to remove a worksheet"
        };
        result.WorkflowHint = "Worksheets listed. Next, read, create, or delete as needed.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ReadWorksheet(SheetCommands commands, string filePath, string? sheetName, string? range)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for read action");

        var result = commands.Read(filePath, sheetName, range ?? "");

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the worksheet name is correct",
                "Verify the range is valid (e.g., 'A1:D10')",
                "Use 'list' to see available worksheets"
            };
            result.WorkflowHint = "Read failed. Ensure the worksheet and range are correct.";
            throw new ModelContextProtocol.McpException($"read failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Process the data as needed",
            "Use 'write' to update worksheet data",
            "Use cell 'get-formula' to inspect formulas"
        };
        result.WorkflowHint = "Data read successfully. Next, process or modify as needed.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string WriteWorksheet(SheetCommands commands, string filePath, string? sheetName, string? dataPath)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(dataPath))
            throw new ModelContextProtocol.McpException("sheetName and range (CSV file path) are required for write action");

        // Read CSV file content before passing to Core command
        if (!File.Exists(dataPath))
            throw new ModelContextProtocol.McpException($"CSV file not found: {dataPath}");

        string csvContent;
        try
        {
            csvContent = File.ReadAllText(dataPath);
        }
        catch (Exception ex)
        {
            throw new ModelContextProtocol.McpException($"Failed to read CSV file '{dataPath}': {ex.Message}");
        }

        var result = commands.Write(filePath, sheetName, csvContent);

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the CSV file exists and is accessible",
                "Verify the worksheet name is correct",
                "Use 'read' to verify written data"
            };
            result.WorkflowHint = "Write failed. Ensure the CSV file and worksheet exist.";
            throw new ModelContextProtocol.McpException($"write failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'read' to verify written data",
            "Use cell 'set-formula' to add formulas",
            "Use PowerQuery to transform data further"
        };
        result.WorkflowHint = "Data written successfully. Next, verify or enhance as needed.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string CreateWorksheet(SheetCommands commands, string filePath, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for create action");

        var result = commands.Create(filePath, sheetName);

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the worksheet name doesn't already exist",
                "Verify the worksheet name is valid",
                "Use 'list' to see existing worksheets"
            };
            result.WorkflowHint = "Create failed. Ensure the worksheet name is unique and valid.";
            throw new ModelContextProtocol.McpException($"create failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'write' to populate the new worksheet",
            "Use 'read' to verify worksheet exists",
            "Use PowerQuery to load data into the sheet"
        };
        result.WorkflowHint = "Worksheet created successfully. Next, populate with data.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string RenameWorksheet(SheetCommands commands, string filePath, string? sheetName, string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ModelContextProtocol.McpException("sheetName and targetName are required for rename action");

        var result = commands.Rename(filePath, sheetName, targetName);

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the source worksheet exists",
                "Verify the target name doesn't already exist",
                "Use 'list' to see available worksheets"
            };
            result.WorkflowHint = "Rename failed. Ensure the source exists and target is unique.";
            throw new ModelContextProtocol.McpException($"rename failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'list' to verify the rename",
            "Use 'read' to access data in the renamed worksheet",
            "Update any formulas referencing the old name"
        };
        result.WorkflowHint = "Worksheet renamed successfully. Next, verify and update references.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string CopyWorksheet(SheetCommands commands, string filePath, string? sheetName, string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ModelContextProtocol.McpException("sheetName and targetName are required for copy action");

        var result = commands.Copy(filePath, sheetName, targetName);

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the source worksheet exists",
                "Verify the target name doesn't already exist",
                "Use 'list' to see available worksheets"
            };
            result.WorkflowHint = "Copy failed. Ensure the source exists and target is unique.";
            throw new ModelContextProtocol.McpException($"copy failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'list' to verify the copy",
            "Use 'read' to access data in the copied worksheet",
            "Modify the copied sheet independently"
        };
        result.WorkflowHint = "Worksheet copied successfully. Next, verify and modify as needed.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteWorksheet(SheetCommands commands, string filePath, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for delete action");

        var result = commands.Delete(filePath, sheetName);

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the worksheet exists",
                "Verify the worksheet is not the only sheet in the workbook",
                "Use 'list' to see available worksheets"
            };
            result.WorkflowHint = "Delete failed. Ensure the worksheet exists and is not the last sheet.";
            throw new ModelContextProtocol.McpException($"delete failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'list' to verify the deletion",
            "Update any formulas referencing the deleted sheet",
            "Review remaining worksheets"
        };
        result.WorkflowHint = "Worksheet deleted successfully. Next, verify and update references.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ClearWorksheet(SheetCommands commands, string filePath, string? sheetName, string? range)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for clear action");

        var result = commands.Clear(filePath, sheetName, range ?? "");

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the worksheet exists",
                "Verify the range is valid (e.g., 'A1:D10')",
                "Use 'list' to see available worksheets"
            };
            result.WorkflowHint = "Clear failed. Ensure the worksheet and range are correct.";
            throw new ModelContextProtocol.McpException($"clear failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'write' to populate the cleared range",
            "Use 'read' to verify the clear operation",
            "Use PowerQuery to reload data"
        };
        result.WorkflowHint = "Range cleared successfully. Next, populate with new data.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string AppendWorksheet(SheetCommands commands, string filePath, string? sheetName, string? dataPath)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(dataPath))
            throw new ModelContextProtocol.McpException("sheetName and range (CSV file path) are required for append action");

        // Read CSV file content before passing to Core command
        if (!File.Exists(dataPath))
            throw new ModelContextProtocol.McpException($"CSV file not found: {dataPath}");

        string csvContent;
        try
        {
            csvContent = File.ReadAllText(dataPath);
        }
        catch (Exception ex)
        {
            throw new ModelContextProtocol.McpException($"Failed to read CSV file '{dataPath}': {ex.Message}");
        }

        var result = commands.Append(filePath, sheetName, csvContent);

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the CSV file exists and is accessible",
                "Verify the worksheet exists",
                "Use 'read' to check existing data before appending"
            };
            result.WorkflowHint = "Append failed. Ensure the CSV file and worksheet exist.";
            throw new ModelContextProtocol.McpException($"append failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'read' to verify appended data",
            "Use cell 'set-formula' to add calculations",
            "Use PowerQuery for further transformations"
        };
        result.WorkflowHint = "Data appended successfully. Next, verify and enhance as needed.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
