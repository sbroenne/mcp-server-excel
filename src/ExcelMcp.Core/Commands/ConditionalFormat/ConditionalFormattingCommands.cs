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

    /// <inheritdoc />
    public ConditionalFormatListResult ListRules(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        var result = new ConditionalFormatListResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? formatConditions = null;

            try
            {
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                range = sheet.Range[rangeAddress];
                formatConditions = range.FormatConditions;

                result.SheetName = sheet.Name;
                result.Rules = ReadFormatConditions(formatConditions);
                result.Success = true;

                return result;
            }
            finally
            {
                ComUtilities.Release(ref formatConditions!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public ConditionalFormatListResult ListWorksheetRules(
        IExcelBatch batch,
        string sheetName)
    {
        var result = new ConditionalFormatListResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = null
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? cells = null;
            dynamic? formatConditions = null;

            try
            {
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                cells = sheet.Cells;
                formatConditions = cells.FormatConditions;

                result.SheetName = sheet.Name;
                result.Rules = ReadFormatConditions(formatConditions);
                result.Success = true;

                return result;
            }
            finally
            {
                ComUtilities.Release(ref formatConditions!);
                ComUtilities.Release(ref cells!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    // === HELPER METHODS ===

    /// <summary>
    /// Reads a FormatConditions collection into a list of rule descriptors.
    /// Each optional COM property read is guarded so unsupported rule types degrade gracefully.
    /// </summary>
    private static List<ConditionalFormatRuleInfo> ReadFormatConditions(dynamic formatConditions)
    {
        var rules = new List<ConditionalFormatRuleInfo>();

        int count = Convert.ToInt32(formatConditions.Count, System.Globalization.CultureInfo.InvariantCulture);
        for (int i = 1; i <= count; i++)
        {
            dynamic? fc = null;
            dynamic? appliesTo = null;
            dynamic? interior = null;
            dynamic? font = null;
            dynamic? borders = null;
            dynamic? edgeBorder = null;

            try
            {
                fc = formatConditions.Item(i);

                var rule = new ConditionalFormatRuleInfo
                {
                    Type = ReadRuleType(fc)
                };

                rule.Operator = ReadRuleOperator(fc);
                rule.Formula1 = ReadRuleString(fc, "Formula1");
                rule.Formula2 = ReadRuleString(fc, "Formula2");
                rule.Priority = ReadRuleInt(fc, "Priority");
                rule.StopIfTrue = ReadRuleBool(fc, "StopIfTrue");

                try
                {
                    appliesTo = fc.AppliesTo;
                    rule.AppliesTo = appliesTo?.Address;
                }
                catch (Exception ex) when (IsComOrBinderException(ex)) { }

                // Interior (fill)
                try
                {
                    interior = fc.Interior;
                    int colorIndex = Convert.ToInt32(interior.ColorIndex, System.Globalization.CultureInfo.InvariantCulture);
                    if (colorIndex != -4142 && colorIndex != -4105) // not None/Automatic
                    {
                        rule.InteriorColor = FormattingHelpers.ColorToHex(Convert.ToInt32(interior.Color, System.Globalization.CultureInfo.InvariantCulture));
                        try { rule.InteriorPattern = Convert.ToInt32(interior.Pattern, System.Globalization.CultureInfo.InvariantCulture); }
                        catch (Exception ex) when (IsComOrBinderException(ex)) { }
                    }
                }
                catch (Exception ex) when (IsComOrBinderException(ex)) { }

                // Font
                try
                {
                    font = fc.Font;
                    int fontColorIndex = Convert.ToInt32(font.ColorIndex, System.Globalization.CultureInfo.InvariantCulture);
                    if (fontColorIndex != -4105 && fontColorIndex != -4142) // not Automatic/None
                    {
                        try { rule.FontColor = FormattingHelpers.ColorToHex(Convert.ToInt32(font.Color, System.Globalization.CultureInfo.InvariantCulture)); }
                        catch (Exception ex) when (IsComOrBinderException(ex)) { }
                    }
                    rule.FontBold = ReadRuleBool(font, "Bold");
                    rule.FontItalic = ReadRuleBool(font, "Italic");
                }
                catch (Exception ex) when (IsComOrBinderException(ex)) { }

                // Borders: scan all four edges and use the first that has a
                // style (rules typically apply borders uniformly, but external
                // rules may set only some edges).
                try
                {
                    borders = fc.Borders;
                    foreach (int edgeIndex in new[] { 7, 8, 9, 10 }) // left, top, bottom, right
                    {
                        edgeBorder = borders.Item(edgeIndex);
                        int lineStyle = Convert.ToInt32(edgeBorder.LineStyle, System.Globalization.CultureInfo.InvariantCulture);
                        if (lineStyle != -4142) // xlLineStyleNone
                        {
                            rule.BorderStyle = BorderStyleToString(lineStyle) ?? lineStyle.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            try { rule.BorderColor = FormattingHelpers.ColorToHex(Convert.ToInt32(edgeBorder.Color, System.Globalization.CultureInfo.InvariantCulture)); }
                            catch (Exception ex) when (IsComOrBinderException(ex)) { }
                            ComUtilities.Release(ref edgeBorder!);
                            break;
                        }
                        ComUtilities.Release(ref edgeBorder!);
                    }
                }
                catch (Exception ex) when (IsComOrBinderException(ex)) { }

                rules.Add(rule);
            }
            finally
            {
                ComUtilities.Release(ref edgeBorder!);
                ComUtilities.Release(ref borders!);
                ComUtilities.Release(ref font!);
                ComUtilities.Release(ref interior!);
                ComUtilities.Release(ref appliesTo!);
                ComUtilities.Release(ref fc!);
            }
        }

        return rules;
    }

    private static string ReadRuleType(dynamic fc)
    {
        try { return ConditionalFormattingTypeToString(Convert.ToInt32(fc.Type, System.Globalization.CultureInfo.InvariantCulture)); }
        catch (Exception ex) when (IsComOrBinderException(ex)) { return "unknown"; }
    }

    private static string? ReadRuleOperator(dynamic fc)
    {
        try { return ConditionalFormattingOperatorToString(Convert.ToInt32(fc.Operator, System.Globalization.CultureInfo.InvariantCulture)); }
        catch (Exception ex) when (IsComOrBinderException(ex)) { return null; }
    }

    private static string? ReadRuleString(dynamic fc, string property)
    {
        try
        {
            string? value = property switch
            {
                "Formula1" => fc.Formula1,
                "Formula2" => fc.Formula2,
                _ => null
            };
            return string.IsNullOrEmpty(value) ? null : value;
        }
        catch (Exception ex) when (IsComOrBinderException(ex))
        {
            return null;
        }
    }

    private static int? ReadRuleInt(dynamic fc, string property)
    {
        try
        {
            var value = property switch
            {
                "Priority" => (object)fc.Priority,
                _ => null
            };
            return value == null ? null : (int?)Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (IsComOrBinderException(ex))
        {
            return null;
        }
    }

    private static bool? ReadRuleBool(dynamic obj, string property)
    {
        try
        {
            var value = property switch
            {
                "StopIfTrue" => (object?)obj.StopIfTrue,
                "Bold" => obj.Bold,
                "Italic" => obj.Italic,
                _ => null
            };
            return value == null ? null : (bool?)Convert.ToBoolean(value, System.Globalization.CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (IsComOrBinderException(ex))
        {
            return null;
        }
    }

    private static bool IsComOrBinderException(Exception ex) =>
        ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException
        or System.Runtime.InteropServices.COMException
        or System.InvalidCastException;

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

    // === REVERSE MAPPINGS (int -> string) for reading existing rules ===

    private static string ConditionalFormattingTypeToString(int type)
    {
        return type switch
        {
            1 => "cellValue", // xlCellValue
            2 => "expression", // xlExpression
            3 => "colorScale", // xlColorScale
            4 => "dataBar", // xlDatabar
            5 => "top10", // xlTop10
            6 => "iconSet", // xlIconSet
            8 => "uniqueValues", // xlUniqueValues
            10 => "blanksCondition", // xlBlanksCondition
            11 => "timePeriod", // xlTimePeriod
            12 => "aboveAverage", // xlAboveAverageCondition
            _ => $"unknown({type})"
        };
    }

    private static string? ConditionalFormattingOperatorToString(int operatorType)
    {
        return operatorType switch
        {
            0 => null, // xlNoOperator (rule type does not use an operator)
            1 => "between", // xlBetween
            2 => "notBetween", // xlNotBetween
            3 => "equal", // xlEqual
            4 => "notEqual", // xlNotEqual
            5 => "greater", // xlGreater
            6 => "less", // xlLess
            7 => "greaterEqual", // xlGreaterEqual
            8 => "lessEqual", // xlLessEqual
            _ => null
        };
    }

    private static string? BorderStyleToString(int lineStyle)
    {
        return lineStyle switch
        {
            -4142 => "none", // xlLineStyleNone
            1 => "continuous", // xlContinuous
            -4115 => "dash", // xlDash
            4 => "dashDot", // xlDashDot
            5 => "dashDotDot", // xlDashDotDot
            -4118 => "dot", // xlDot
            -4119 => "double", // xlDouble
            13 => "slantDashDot", // xlSlantDashDot
            _ => null
        };
    }
}



