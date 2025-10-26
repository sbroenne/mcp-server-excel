using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel cell manipulation tool for MCP server.
/// Handles individual cell operations for precise data control.
///
/// LLM Usage Patterns:
/// - Use "get-value" to read individual cell contents
/// - Use "set-value" to write data to specific cells
/// - Use "get-formula" to examine cell formulas
/// - Use "set-formula" to create calculated cells
///
/// Note: For bulk operations, use ExcelWorksheetTool instead.
/// This tool is optimized for precise, single-cell operations.
/// </summary>
[McpServerToolType]
public static class ExcelCellTool
{
    /// <summary>
    /// Manage individual Excel cells - values and formulas for precise control
    /// </summary>
    [McpServerTool(Name = "excel_cell")]
    [Description("Manage individual Excel cell values and formulas. Supports: get-value, set-value, get-formula, set-formula. Optional batchId for batch sessions.")]
    public static async Task<string> ExcelCell(
        [Required]
        [RegularExpression("^(get-value|set-value|get-formula|set-formula)$")]
        [Description("Action: get-value, set-value, get-formula, set-formula")]
        string action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Worksheet name")]
        string sheetName,

        [Required]
        [RegularExpression(@"^[A-Z]+[0-9]+$")]
        [Description("Cell address (e.g., 'A1', 'B5')")]
        string cellAddress,

        [StringLength(32767)]
        [Description("Value or formula to set (for set-value/set-formula actions)")]
        string? value = null,
        
        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var cellCommands = new CellCommands();

            switch (action.ToLowerInvariant())
            {
                case "get-value":
                    return await GetCellValueAsync(cellCommands, excelPath, sheetName, cellAddress, batchId);
                case "set-value":
                    return await SetCellValueAsync(cellCommands, excelPath, sheetName, cellAddress, value, batchId);
                case "get-formula":
                    return await GetCellFormulaAsync(cellCommands, excelPath, sheetName, cellAddress, batchId);
                case "set-formula":
                    return await SetCellFormulaAsync(cellCommands, excelPath, sheetName, cellAddress, value, batchId);
                default:
                    ExcelToolsBase.ThrowUnknownAction(action, "get-value", "set-value", "get-formula", "set-formula");
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

    private static async Task<string> GetCellValueAsync(CellCommands commands, string excelPath, string sheetName, string cellAddress, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.GetValueAsync(batch, sheetName, cellAddress));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the worksheet and cell address are correct",
                "Use worksheet 'list' to verify worksheet exists",
                "Verify the cell address format (e.g., 'A1', 'B5')"
            };
            result.WorkflowHint = "Get-value failed. Ensure the worksheet and cell exist.";
            throw new ModelContextProtocol.McpException($"get-value failed for '{excelPath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'set-value' to update the cell",
            "Use 'get-formula' to check if cell has a formula",
            "Use worksheet 'read' for multiple cells"
        };
        result.WorkflowHint = "Cell value retrieved. Next, update or inspect formula.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetCellValueAsync(CellCommands commands, string excelPath, string sheetName, string cellAddress, string? value, string? batchId)
    {
        if (value == null)
            throw new ModelContextProtocol.McpException("value is required for set-value action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.SetValueAsync(batch, sheetName, cellAddress, value));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the worksheet and cell address are correct",
                "Verify the value format is appropriate",
                "Use worksheet 'write' for bulk data updates"
            };
            result.WorkflowHint = "Set-value failed. Ensure the worksheet and cell are valid.";
            throw new ModelContextProtocol.McpException($"set-value failed for '{excelPath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'get-value' to verify the update",
            "Use 'set-formula' to add calculations",
            "Use worksheet 'read' to view surrounding cells"
        };
        result.WorkflowHint = "Cell value set. Next, verify or add formulas.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetCellFormulaAsync(CellCommands commands, string excelPath, string sheetName, string cellAddress, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.GetFormulaAsync(batch, sheetName, cellAddress));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the worksheet and cell address are correct",
                "Use 'get-value' to retrieve the cell value instead",
                "Verify the cell address format (e.g., 'A1', 'B5')"
            };
            result.WorkflowHint = "Get-formula failed. Ensure the worksheet and cell exist.";
            throw new ModelContextProtocol.McpException($"get-formula failed for '{excelPath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'set-formula' to update the formula",
            "Use 'get-value' to see the calculated result",
            "Analyze the formula for optimization"
        };
        result.WorkflowHint = "Cell formula retrieved. Next, update or analyze.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetCellFormulaAsync(CellCommands commands, string excelPath, string sheetName, string cellAddress, string? value, string? batchId)
    {
        if (string.IsNullOrEmpty(value))
            throw new ModelContextProtocol.McpException("value (formula) is required for set-formula action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.SetFormulaAsync(batch, sheetName, cellAddress, value));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the formula syntax is correct",
                "Verify all cell references in the formula exist",
                "Use 'get-value' to see if formula calculated correctly"
            };
            result.WorkflowHint = "Set-formula failed. Ensure the formula syntax is valid.";
            throw new ModelContextProtocol.McpException($"set-formula failed for '{excelPath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'get-value' to verify the calculated result",
            "Use 'get-formula' to confirm the formula was set",
            "Test the formula with different input values"
        };
        result.WorkflowHint = "Formula set. Next, verify calculation and test.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
