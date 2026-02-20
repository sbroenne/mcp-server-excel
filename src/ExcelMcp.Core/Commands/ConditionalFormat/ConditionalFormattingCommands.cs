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
                    : ctx.Book.Worksheets[sheetName];

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
                    : ctx.Book.Worksheets[sheetName];

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
            "cell-value" => 1, // xlCellValue (kebab-case alias)
            "expression" => 2, // xlExpression
            "colorscale" => 3, // xlColorScale
            "color-scale" => 3, // xlColorScale (kebab-case alias)
            "databar" => 4, // xlDatabar
            "data-bar" => 4, // xlDatabar (kebab-case alias)
            "top10" => 5, // xlTop10
            "iconset" => 6, // xlIconSet
            "icon-set" => 6, // xlIconSet (kebab-case alias)
            "uniquevalues" => 8, // xlUniqueValues
            "unique-values" => 8, // xlUniqueValues (kebab-case alias)
            "blankscondition" => 10, // xlBlanksCondition
            "blanks-condition" => 10, // xlBlanksCondition (kebab-case alias)
            "timeperiod" => 11, // xlTimePeriod
            "time-period" => 11, // xlTimePeriod (kebab-case alias)
            "aboveaverage" => 12, // xlAboveAverageCondition
            "above-average" => 12, // xlAboveAverageCondition (kebab-case alias)
            _ => throw new ArgumentException(
                $"Invalid conditional formatting type: '{type}'. " +
                "Valid values: cellValue, expression, colorScale, dataBar, top10, iconSet, uniqueValues, blanksCondition, timePeriod, aboveAverage")
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
            "not-between" => 2, // xlNotBetween (kebab-case alias)
            "equal" => 3, // xlEqual
            "notequal" => 4, // xlNotEqual
            "not-equal" => 4, // xlNotEqual (kebab-case alias)
            "greater" => 5, // xlGreater
            "greaterthan" => 5, // xlGreater (alias)
            "less" => 6, // xlLess
            "lessthan" => 6, // xlLess (alias)
            "greaterequal" => 7, // xlGreaterEqual
            "greater-equal" => 7, // xlGreaterEqual (kebab-case alias)
            "greaterthanorequal" => 7, // xlGreaterEqual (alias)
            ">=" => 7, // xlGreaterEqual (symbol alias)
            "lessequal" => 8, // xlLessEqual
            "less-equal" => 8, // xlLessEqual (kebab-case alias)
            "lessthanorequal" => 8, // xlLessEqual (alias)
            "<=" => 8, // xlLessEqual (symbol alias)
            "=" => 3, // xlEqual (symbol alias)
            "<>" => 4, // xlNotEqual (symbol alias)
            ">" => 5, // xlGreater (symbol alias)
            "<" => 6, // xlLess (symbol alias)
            _ => throw new ArgumentException($"Unknown operator type: '{operatorType}'. Valid values: between, notBetween, equal, notEqual, greater, less, greaterEqual, lessEqual")
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



