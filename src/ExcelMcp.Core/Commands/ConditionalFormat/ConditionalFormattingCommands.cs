using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Implementation of conditional formatting commands
/// </summary>
public partial class ConditionalFormattingCommands : IConditionalFormattingCommands
{
    /// <inheritdoc />
    public OperationResult AddRule(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string ruleType,
        string? operatorType,
        string? formula1,
        string? formula2,
        string? interiorColor = null,
        string? interiorPattern = null,
        string? fontColor = null,
        bool? fontBold = null,
        bool? fontItalic = null,
        string? borderStyle = null,
        string? borderColor = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? formatConditions = null;
            dynamic? formatCondition = null;
            dynamic? interior = null;
            dynamic? font = null;
            dynamic? borders = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get format conditions
                formatConditions = range.FormatConditions;

                // Parse rule type and operator
                var xlType = ParseConditionalFormattingType(ruleType);
                var xlOperator = ParseConditionalFormattingOperator(operatorType);

                // Add format condition
                formatCondition = formatConditions.Add(
                    Type: xlType,
                    Operator: xlOperator,
                    Formula1: formula1 ?? "",
                    Formula2: formula2 ?? "");

                // Apply Interior formatting
                if (!string.IsNullOrEmpty(interiorColor) || !string.IsNullOrEmpty(interiorPattern))
                {
                    interior = formatCondition.Interior;
                    if (!string.IsNullOrEmpty(interiorColor))
                        interior.Color = FormattingHelpers.ParseColor(interiorColor);
                    if (!string.IsNullOrEmpty(interiorPattern))
                        interior.Pattern = ParseInteriorPattern(interiorPattern);
                }

                // Apply Font formatting
                if (!string.IsNullOrEmpty(fontColor) || fontBold.HasValue || fontItalic.HasValue)
                {
                    font = formatCondition.Font;
                    if (!string.IsNullOrEmpty(fontColor))
                        font.Color = FormattingHelpers.ParseColor(fontColor);
                    if (fontBold.HasValue)
                        font.Bold = fontBold.Value;
                    if (fontItalic.HasValue)
                        font.Italic = fontItalic.Value;
                }

                // Apply Border formatting
                if (!string.IsNullOrEmpty(borderStyle) || !string.IsNullOrEmpty(borderColor))
                {
                    borders = formatCondition.Borders;
                    if (!string.IsNullOrEmpty(borderStyle))
                    {
                        var xlBorderStyle = FormattingHelpers.ParseBorderStyle(borderStyle);
                        // Apply to all four borders
                        borders.Item(7).LineStyle = xlBorderStyle;  // xlEdgeLeft
                        borders.Item(8).LineStyle = xlBorderStyle;  // xlEdgeTop
                        borders.Item(9).LineStyle = xlBorderStyle;  // xlEdgeBottom
                        borders.Item(10).LineStyle = xlBorderStyle; // xlEdgeRight
                    }
                    if (!string.IsNullOrEmpty(borderColor))
                    {
                        var color = FormattingHelpers.ParseColor(borderColor);
                        borders.Item(7).Color = color;  // xlEdgeLeft
                        borders.Item(8).Color = color;  // xlEdgeTop
                        borders.Item(9).Color = color;  // xlEdgeBottom
                        borders.Item(10).Color = color; // xlEdgeRight
                    }
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref borders!);
                ComUtilities.Release(ref font!);
                ComUtilities.Release(ref interior!);
                ComUtilities.Release(ref formatCondition!);
                ComUtilities.Release(ref formatConditions!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ClearRules(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? formatConditions = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get and delete format conditions
                formatConditions = range.FormatConditions;
                formatConditions.Delete();

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref formatConditions!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    // === HELPER METHODS ===

    private static int ParseConditionalFormattingType(string type)
    {
        return type.ToLowerInvariant() switch
        {
            "cellvalue" => 1, // xlCellValue
            "expression" => 2, // xlExpression
            "colorscale" => 3, // xlColorScale
            "databar" => 4, // xlDatabar
            "top10" => 5, // xlTop10
            "iconset" => 6, // xlIconSet
            "uniquevalues" => 8, // xlUniqueValues
            "blankscondition" => 10, // xlBlanksCondition
            "timeperiod" => 11, // xlTimePeriod
            "aboveaverage" => 12, // xlAboveAverageCondition
            _ => throw new ArgumentException($"Invalid conditional formatting type: {type}")
        };
    }

    private static int ParseConditionalFormattingOperator(string? operatorType)
    {
        if (string.IsNullOrEmpty(operatorType))
            return 3; // xlEqual (default)

        return operatorType.ToLowerInvariant() switch
        {
            "between" => 1, // xlBetween
            "notbetween" => 2, // xlNotBetween
            "equal" => 3, // xlEqual
            "notequal" => 4, // xlNotEqual
            "greater" => 5, // xlGreater
            "less" => 6, // xlLess
            "greaterequal" => 7, // xlGreaterEqual
            "lessequal" => 8, // xlLessEqual
            _ => throw new ArgumentException($"Unknown operator type: {operatorType}. Valid values: between, notBetween, equal, notEqual, greater, less, greaterEqual, lessEqual")
        };
    }

    private static int ParseInteriorPattern(string pattern)
    {
        if (int.TryParse(pattern, out var patternValue))
            return patternValue;

        return pattern.ToLowerInvariant() switch
        {
            "none" => -4142, // xlPatternNone
            "solid" => 1, // xlPatternSolid
            "gray50" => 9, // xlPatternGray50
            "gray75" => 10, // xlPatternGray75
            "gray25" => 11, // xlPatternGray25
            _ => throw new ArgumentException($"Unknown interior pattern: {pattern}. Use pattern constant or: none, solid, gray50, gray75, gray25")
        };
    }
}



