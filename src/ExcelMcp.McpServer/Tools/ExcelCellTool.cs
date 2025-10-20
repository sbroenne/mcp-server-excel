using Sbroenne.ExcelMcp.Core.Commands;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;

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
    [Description("Manage individual Excel cell values and formulas. Supports: get-value, set-value, get-formula, set-formula.")]
    public static string ExcelCell(
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
        string? value = null)
    {
        try
        {
            var cellCommands = new CellCommands();

            switch (action.ToLowerInvariant())
            {
                case "get-value":
                    return GetCellValue(cellCommands, excelPath, sheetName, cellAddress);
                case "set-value":
                    return SetCellValue(cellCommands, excelPath, sheetName, cellAddress, value);
                case "get-formula":
                    return GetCellFormula(cellCommands, excelPath, sheetName, cellAddress);
                case "set-formula":
                    return SetCellFormula(cellCommands, excelPath, sheetName, cellAddress, value);
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

    private static string GetCellValue(CellCommands commands, string excelPath, string sheetName, string cellAddress)
    {
        var result = commands.GetValue(excelPath, sheetName, cellAddress);
        
        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-value failed for '{excelPath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetCellValue(CellCommands commands, string excelPath, string sheetName, string cellAddress, string? value)
    {
        if (value == null)
            throw new ModelContextProtocol.McpException("value is required for set-value action");

        var result = commands.SetValue(excelPath, sheetName, cellAddress, value);
        
        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-value failed for '{excelPath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string GetCellFormula(CellCommands commands, string excelPath, string sheetName, string cellAddress)
    {
        var result = commands.GetFormula(excelPath, sheetName, cellAddress);
        
        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-formula failed for '{excelPath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetCellFormula(CellCommands commands, string excelPath, string sheetName, string cellAddress, string? value)
    {
        if (string.IsNullOrEmpty(value))
            throw new ModelContextProtocol.McpException("value (formula) is required for set-formula action");

        var result = commands.SetFormula(excelPath, sheetName, cellAddress, value);
        
        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-formula failed for '{excelPath}': {result.ErrorMessage}");
        }
        
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}