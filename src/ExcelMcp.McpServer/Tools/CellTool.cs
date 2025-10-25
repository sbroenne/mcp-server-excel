using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel cell manipulation tool for MCP server.
/// Handles individual cell operations and formatting for precise control.
///
/// LLM Usage Patterns:
/// - Use "get-value" to read individual cell contents
/// - Use "set-value" to write data to specific cells
/// - Use "get-formula" to examine cell formulas
/// - Use "set-formula" to create calculated cells
/// - Use "set-background-color" to apply background colors
/// - Use "set-font-color" to change text colors
/// - Use "set-font" to configure font properties
/// - Use "set-border" to add borders
/// - Use "set-number-format" to format numbers/dates
/// - Use "set-alignment" to align cell content
/// - Use "clear-formatting" to remove all formatting
///
/// Note: For bulk operations, use ExcelWorksheetTool instead.
/// This tool is optimized for precise, single-cell operations.
/// </summary>
[McpServerToolType]
public static class CellTool
{
    /// <summary>
    /// Manage individual Excel cells - values, formulas, and formatting for precise control
    /// </summary>
    [McpServerTool(Name = "cell")]
    [Description("Manage individual Excel cell values, formulas, and formatting. Supports: get-value, set-value, get-formula, set-formula, set-background-color, set-font-color, set-font, set-border, set-number-format, set-alignment, clear-formatting.")]
    public static string Cell(
        [Required]
        [RegularExpression("^(get-value|set-value|get-formula|set-formula|set-background-color|set-font-color|set-font|set-border|set-number-format|set-alignment|clear-formatting)$")]
        [Description("Action: get-value, set-value, get-formula, set-formula, set-background-color, set-font-color, set-font, set-border, set-number-format, set-alignment, clear-formatting")]
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
        [Description("Cell address or range (e.g., 'A1', 'B5', 'A1:D10')")]
        string cellAddress,

        [StringLength(32767)]
        [Description("Value or formula to set (for set-value/set-formula actions)")]
        string? value = null,

        [StringLength(255)]
        [Description("Color in hex (#RRGGBB), RGB (r,g,b), or color number (for color actions)")]
        string? color = null,

        [StringLength(255)]
        [Description("Font name (for set-font action)")]
        string? fontName = null,

        [Range(1, 409)]
        [Description("Font size (for set-font action)")]
        int? fontSize = null,

        [Description("Bold (for set-font action)")]
        bool? bold = null,

        [Description("Italic (for set-font action)")]
        bool? italic = null,

        [Description("Underline (for set-font action)")]
        bool? underline = null,

        [StringLength(50)]
        [Description("Border style: thin, dash, dot, double, none (for set-border action)")]
        string? borderStyle = null,

        [StringLength(255)]
        [Description("Border color (for set-border action)")]
        string? borderColor = null,

        [StringLength(255)]
        [Description("Number format string (for set-number-format action)")]
        string? format = null,

        [StringLength(50)]
        [Description("Horizontal alignment: left, center, right, justify (for set-alignment action)")]
        string? horizontal = null,

        [StringLength(50)]
        [Description("Vertical alignment: top, center, bottom (for set-alignment action)")]
        string? vertical = null,

        [Description("Wrap text (for set-alignment action)")]
        bool? wrapText = null)
    {
        try
        {
            var cellCommands = new CellCommands();

            return action.ToLowerInvariant() switch
            {
                "get-value" => GetCellValue(cellCommands, excelPath, sheetName, cellAddress),
                "set-value" => SetCellValue(cellCommands, excelPath, sheetName, cellAddress, value),
                "get-formula" => GetCellFormula(cellCommands, excelPath, sheetName, cellAddress),
                "set-formula" => SetCellFormula(cellCommands, excelPath, sheetName, cellAddress, value),
                "set-background-color" => SetBackgroundColor(cellCommands, excelPath, sheetName, cellAddress, color),
                "set-font-color" => SetFontColor(cellCommands, excelPath, sheetName, cellAddress, color),
                "set-font" => SetFont(cellCommands, excelPath, sheetName, cellAddress, fontName, fontSize, bold, italic, underline),
                "set-border" => SetBorder(cellCommands, excelPath, sheetName, cellAddress, borderStyle, borderColor),
                "set-number-format" => SetNumberFormat(cellCommands, excelPath, sheetName, cellAddress, format),
                "set-alignment" => SetAlignment(cellCommands, excelPath, sheetName, cellAddress, horizontal, vertical, wrapText),
                "clear-formatting" => ClearFormatting(cellCommands, excelPath, sheetName, cellAddress),
                _ => throw new InvalidOperationException($"Unknown action: {action}")
            };
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-value failed for '{excelPath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string GetCellFormula(CellCommands commands, string excelPath, string sheetName, string cellAddress)
    {
        var result = commands.GetFormula(excelPath, sheetName, cellAddress);

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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-formula failed for '{excelPath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetBackgroundColor(CellCommands commands, string excelPath, string sheetName, string cellAddress, string? color)
    {
        if (string.IsNullOrEmpty(color))
            throw new ModelContextProtocol.McpException("color is required for set-background-color action");

        var result = commands.SetBackgroundColor(excelPath, sheetName, cellAddress, color);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-background-color failed for '{excelPath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetFontColor(CellCommands commands, string excelPath, string sheetName, string cellAddress, string? color)
    {
        if (string.IsNullOrEmpty(color))
            throw new ModelContextProtocol.McpException("color is required for set-font-color action");

        var result = commands.SetFontColor(excelPath, sheetName, cellAddress, color);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-font-color failed for '{excelPath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetFont(CellCommands commands, string excelPath, string sheetName, string cellAddress, 
        string? fontName, int? fontSize, bool? bold, bool? italic, bool? underline)
    {
        var result = commands.SetFont(excelPath, sheetName, cellAddress, fontName, fontSize, bold, italic, underline);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-font failed for '{excelPath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetBorder(CellCommands commands, string excelPath, string sheetName, string cellAddress, 
        string? borderStyle, string? borderColor)
    {
        if (string.IsNullOrEmpty(borderStyle))
            throw new ModelContextProtocol.McpException("borderStyle is required for set-border action");

        var result = commands.SetBorder(excelPath, sheetName, cellAddress, borderStyle, borderColor);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-border failed for '{excelPath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetNumberFormat(CellCommands commands, string excelPath, string sheetName, string cellAddress, string? format)
    {
        if (string.IsNullOrEmpty(format))
            throw new ModelContextProtocol.McpException("format is required for set-number-format action");

        var result = commands.SetNumberFormat(excelPath, sheetName, cellAddress, format);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-number-format failed for '{excelPath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetAlignment(CellCommands commands, string excelPath, string sheetName, string cellAddress, 
        string? horizontal, string? vertical, bool? wrapText)
    {
        var result = commands.SetAlignment(excelPath, sheetName, cellAddress, horizontal, vertical, wrapText);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-alignment failed for '{excelPath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ClearFormatting(CellCommands commands, string excelPath, string sheetName, string cellAddress)
    {
        var result = commands.ClearFormatting(excelPath, sheetName, cellAddress);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"clear-formatting failed for '{excelPath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
