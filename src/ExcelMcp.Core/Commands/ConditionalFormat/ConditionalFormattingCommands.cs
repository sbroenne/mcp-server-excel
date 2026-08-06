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
        string? borderColor = null,
        string? colorScaleMinType = null,
        string? colorScaleMinValue = null,
        string? colorScaleMinColor = null,
        string? colorScaleMidType = null,
        string? colorScaleMidValue = null,
        string? colorScaleMidColor = null,
        string? colorScaleMaxType = null,
        string? colorScaleMaxValue = null,
        string? colorScaleMaxColor = null,
        string? dataBarColor = null,
        string? dataBarNegativeColor = null,
        string? dataBarDirection = null,
        bool? dataBarShowValue = null,
        string? dataBarMinType = null,
        string? dataBarMinValue = null,
        string? dataBarMaxType = null,
        string? dataBarMaxValue = null,
        string? iconSetId = null,
        bool? iconSetReverse = null,
        bool? iconSetShowIconOnly = null,
        string? iconThreshold1Type = null,
        string? iconThreshold1Value = null,
        string? iconThreshold2Type = null,
        string? iconThreshold2Value = null,
        string? iconThreshold3Type = null,
        string? iconThreshold3Value = null,
        string? iconThreshold4Type = null,
        string? iconThreshold4Value = null,
        int? rank = null,
        bool? top10Percent = null,
        string? topBottom = null,
        string? aboveBelow = null,
        string? datePeriod = null)
    {
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

                var normalizedType = NormalizeRuleType(ruleType);

                switch (normalizedType)
                {
                    case "cellvalue":
                    case "expression":
                        AddBasicRule(formatConditions, normalizedType, operatorType, formula1, formula2,
                            interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                        break;

                    case "colorscale":
                        AddColorScaleRule(formatConditions,
                            colorScaleMinType, colorScaleMinValue, colorScaleMinColor,
                            colorScaleMidType, colorScaleMidValue, colorScaleMidColor,
                            colorScaleMaxType, colorScaleMaxValue, colorScaleMaxColor);
                        break;

                    case "databar":
                        AddDataBarRule(formatConditions,
                            dataBarColor, dataBarNegativeColor, dataBarDirection, dataBarShowValue,
                            dataBarMinType, dataBarMinValue, dataBarMaxType, dataBarMaxValue);
                        break;

                    case "iconset":
                        AddIconSetRule(ctx.Book, formatConditions,
                            iconSetId, iconSetReverse, iconSetShowIconOnly,
                            new[]
                            {
                                (iconThreshold1Type, iconThreshold1Value),
                                (iconThreshold2Type, iconThreshold2Value),
                                (iconThreshold3Type, iconThreshold3Value),
                                (iconThreshold4Type, iconThreshold4Value)
                            });
                        break;

                    case "top10":
                        AddTop10Rule(formatConditions, rank, top10Percent, topBottom,
                            interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                        break;

                    case "aboveaverage":
                        AddAboveAverageRule(formatConditions, aboveBelow,
                            interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                        break;

                    case "uniquevalues":
                        AddUniqueValuesRule(formatConditions, false,
                            interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                        break;

                    case "timeperiod":
                        AddTimePeriodRule(formatConditions, datePeriod,
                            interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                        break;

                    case "blankscondition":
                        AddSimpleRule(formatConditions, 10 /* xlBlanksCondition */,
                            interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                        break;

                    default:
                        throw new ArgumentException(
                            $"Invalid conditional formatting type: '{ruleType}'. " +
                            "Valid values: cellValue, expression, colorScale, dataBar, top10, iconSet, uniqueValues, blanksCondition, timePeriod, aboveAverage");
                }

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

    /// <summary>
    /// Adds a basic (cellValue/expression) rule and applies interior/font/border formatting.
    /// </summary>
    private static void AddBasicRule(
        dynamic formatConditions,
        string normalizedType,
        string? operatorType,
        string? formula1,
        string? formula2,
        string? interiorColor,
        string? interiorPattern,
        string? fontColor,
        bool? fontBold,
        bool? fontItalic,
        string? borderStyle,
        string? borderColor)
    {
        dynamic? formatCondition = null;

        try
        {
            var xlType = normalizedType == "expression" ? 2 : 1;
            var xlOperator = ValidateAndParseBasicRuleArguments(
                normalizedType,
                operatorType,
                formula1,
                formula2);

            formatCondition = formatConditions.Add(
                Type: xlType,
                Operator: xlOperator,
                Formula1: formula1 ?? "",
                Formula2: formula2 ?? "");

            ApplyRuleFormatting(formatCondition,
                interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
        }
        finally
        {
            ComUtilities.Release(ref formatCondition!);
        }
    }

    /// <summary>
    /// Applies interior/font/border formatting to a format condition. Shared by rule types
    /// that support cell formatting (cellValue, expression, top10, aboveAverage, timePeriod, etc.).
    /// </summary>
    private static void ApplyRuleFormatting(
        dynamic formatCondition,
        string? interiorColor,
        string? interiorPattern,
        string? fontColor,
        bool? fontBold,
        bool? fontItalic,
        string? borderStyle,
        string? borderColor)
    {
        dynamic? interior = null;
        dynamic? font = null;
        dynamic? borders = null;

        try
        {
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
                // NOTE: FormatCondition.Borders is a 4-item collection indexed 1-4
                // (left/top/bottom/right), unlike Range.Borders which uses the
                // xlEdgeLeft(7)/xlEdgeTop(8)/xlEdgeBottom(9)/xlEdgeRight(10) constants.
                // Item(7-10) on FormatCondition.Borders returns an unbound placeholder
                // that can be read but throws COMException when its properties are set.
                var xlBorderStyle = !string.IsNullOrEmpty(borderStyle)
                    ? FormattingHelpers.ParseBorderStyle(borderStyle)
                    : (int?)null;
                var color = !string.IsNullOrEmpty(borderColor)
                    ? FormattingHelpers.ParseColor(borderColor)
                    : (int?)null;

                for (var borderIndex = 1; borderIndex <= 4; borderIndex++)
                {
                    dynamic? border = null;
                    try
                    {
                        border = borders.Item(borderIndex);
                        if (xlBorderStyle.HasValue)
                            border.LineStyle = xlBorderStyle.Value;
                        if (color.HasValue)
                            border.Color = color.Value;
                    }
                    finally
                    {
                        ComUtilities.Release(ref border!);
                    }
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref borders!);
            ComUtilities.Release(ref font!);
            ComUtilities.Release(ref interior!);
        }
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

                int typeNum = ReadRuleTypeInt(fc);
                var rule = new ConditionalFormatRuleInfo
                {
                    Type = typeNum >= 0 ? ConditionalFormattingTypeToString(typeNum) : "unknown"
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
                    // NOTE: FormatCondition.Borders is a 4-item collection indexed 1-4
                    // (left/top/bottom/right) - see write-side note in AddRule above.
                    foreach (int edgeIndex in new[] { 1, 2, 3, 4 }) // left, top, bottom, right
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

                // Type-specific configuration (visual rule types).
                switch (typeNum)
                {
                    case 3: rule.ColorScaleCriteria = ReadColorScale(fc); break;   // xlColorScale
                    case 4: rule.DataBar = ReadDataBar(fc); break;                 // xlDatabar
                    case 5: rule.Top10 = ReadTop10(fc); break;                     // xlTop10
                    case 6: rule.IconSet = ReadIconSet(fc); break;                 // xlIconSet
                    case 11: rule.DatePeriod = ReadTimePeriod(fc); break;          // xlTimePeriod
                    case 12: rule.AboveBelow = ReadAboveBelow(fc); break;          // xlAboveAverageCondition
                }

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

    private static int ReadRuleTypeInt(dynamic fc)
    {
        try { return Convert.ToInt32(fc.Type, System.Globalization.CultureInfo.InvariantCulture); }
        catch (Exception ex) when (IsComOrBinderException(ex)) { return -1; }
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

    private static int ValidateAndParseBasicRuleArguments(
        string normalizedType,
        string? operatorType,
        string? formula1,
        string? formula2)
    {
        if (string.IsNullOrWhiteSpace(formula1))
        {
            throw new ArgumentException(
                $"formula1 is required for {normalizedType} conditional formatting rules.",
                nameof(formula1));
        }

        if (normalizedType == "expression")
            return 3; // Operator is ignored by Excel for xlExpression.

        if (string.IsNullOrWhiteSpace(operatorType))
        {
            throw new ArgumentException(
                "operatorType is required for cellValue conditional formatting rules.",
                nameof(operatorType));
        }

        var parsedOperator = ParseConditionalFormattingOperator(operatorType);
        if (parsedOperator is 1 or 2 && string.IsNullOrWhiteSpace(formula2))
        {
            throw new ArgumentException(
                $"formula2 is required when operatorType is '{operatorType}'.",
                nameof(formula2));
        }

        return parsedOperator;
    }

    private static int ParseConditionalFormattingOperator(string? operatorType)
    {
        return operatorType!.ToLowerInvariant() switch
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


