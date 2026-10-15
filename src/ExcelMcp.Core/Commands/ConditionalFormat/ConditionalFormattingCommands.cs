using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Excel = Microsoft.Office.Interop.Excel;

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
        string? datePeriod = null,
        int? priority = null,
        bool? stopIfTrue = null)
    {
        var normalizedType = NormalizeRuleType(ruleType);
        if (priority is < 1)
            throw new ArgumentOutOfRangeException(nameof(priority));
        if (stopIfTrue.HasValue && normalizedType is "colorscale" or "databar" or "iconset")
            throw new ArgumentException("StopIfTrue is not available for color-scale, data-bar, or icon-set rules.", nameof(stopIfTrue));
        string?[] colors = normalizedType switch
        {
            "colorscale" => [colorScaleMinColor, colorScaleMidColor, colorScaleMaxColor],
            "databar" => [dataBarColor, dataBarNegativeColor],
            "cellvalue" or "expression" or "top10" or "aboveaverage" or "uniquevalues"
                or "timeperiod" or "blankscondition" => [interiorColor, fontColor, borderColor],
            _ => []
        };
        foreach (var color in colors)
        {
            if (!string.IsNullOrEmpty(color)) { _ = FormattingHelpers.ParseColor(color); }
        }
        ValidateRuleOptions(normalizedType, interiorPattern, borderStyle,
            [colorScaleMinType, colorScaleMidType, colorScaleMaxType],
            dataBarDirection, [dataBarMinType, dataBarMaxType],
            iconSetId, [iconThreshold1Type, iconThreshold2Type, iconThreshold3Type, iconThreshold4Type],
            rank, top10Percent, topBottom, aboveBelow);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.FormatConditions? formatConditions = null;
            Excel.Range? cells = null;
            Excel.FormatConditions? worksheetConditions = null;
            object? createdRule = null;

            try
            {
                sheet = RangeHelpers.ResolveWorksheet(ctx.Book, sheetName);

                range = sheet.Range[rangeAddress];
                formatConditions = range.FormatConditions;

                if (priority.HasValue)
                {
                    cells = sheet.Cells;
                    worksheetConditions = cells.FormatConditions;
                    int count = Convert.ToInt32(worksheetConditions.Count, System.Globalization.CultureInfo.InvariantCulture);
                    List<ConditionalFormatRuleInfo> existing = ReadFormatConditions(worksheetConditions, ct);
                    int lastPriority = Math.Max(count, existing.Select(item => item.Priority ?? 0).DefaultIfEmpty(0).Max());
                    if (priority.Value > lastPriority + 1)
                        throw new ArgumentOutOfRangeException(nameof(priority), "Priority must not exceed the greatest native worksheet priority plus one.");
                }

                int rulesBefore = Convert.ToInt32(formatConditions.Count, System.Globalization.CultureInfo.InvariantCulture);
                try
                {
                    switch (normalizedType)
                    {
                        case "cellvalue":
                        case "expression":
                            createdRule = AddBasicRule(formatConditions, normalizedType, operatorType, formula1, formula2,
                                interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                            break;

                        case "colorscale":
                            createdRule = AddColorScaleRule(formatConditions,
                                colorScaleMinType, colorScaleMinValue, colorScaleMinColor,
                                colorScaleMidType, colorScaleMidValue, colorScaleMidColor,
                                colorScaleMaxType, colorScaleMaxValue, colorScaleMaxColor);
                            break;

                        case "databar":
                            createdRule = AddDataBarRule(formatConditions,
                                dataBarColor, dataBarNegativeColor, dataBarDirection, dataBarShowValue,
                                dataBarMinType, dataBarMinValue, dataBarMaxType, dataBarMaxValue);
                            break;

                        case "iconset":
                            createdRule = AddIconSetRule(ctx.Book, formatConditions,
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
                            createdRule = AddTop10Rule(formatConditions, rank, top10Percent, topBottom,
                                interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                            break;

                        case "aboveaverage":
                            createdRule = AddAboveAverageRule(formatConditions, aboveBelow,
                                interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                            break;

                        case "uniquevalues":
                            createdRule = AddUniqueValuesRule(formatConditions, false,
                                interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                            break;

                        case "timeperiod":
                            createdRule = AddTimePeriodRule(formatConditions, datePeriod,
                                interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                            break;

                        case "blankscondition":
                            createdRule = AddSimpleRule(formatConditions, 10 /* xlBlanksCondition */,
                                interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                            break;

                        default:
                            throw new ArgumentException(
                                $"Invalid conditional formatting type: '{ruleType}'. " +
                                "Valid values: cellValue, expression, colorScale, dataBar, top10, iconSet, uniqueValues, blanksCondition, timePeriod, aboveAverage");
                    }

                    if (stopIfTrue.HasValue)
                        SetNativeStopIfTrue(createdRule!, stopIfTrue.Value);
                    if (priority.HasValue && GetNativePriority(createdRule!) != priority.Value)
                        SetNativePriority(createdRule!, priority.Value);
                    if (priority.HasValue && GetNativePriority(createdRule!) != priority.Value)
                        throw new InvalidOperationException("Excel did not apply the requested new-rule priority. List worksheet rules before retrying.");
                }
                catch (Exception ex) when (CreatedObjectFailure.CanReport(ex)
                    && Convert.ToInt32(formatConditions.Count, System.Globalization.CultureInfo.InvariantCulture) > rulesBefore)
                {
                    string actualSheet = sheet.Name;
                    throw CreatedObjectFailure.CreateDescribed(
                        $"Setting up the {ruleType} rule",
                        $"a {ruleType} rule on range '{rangeAddress}' of sheet '{actualSheet}'",
                        $"The {ruleType} rule remains on range '{rangeAddress}' of sheet '{actualSheet}'; list the rules there, then fix or remove it if appropriate.",
                        ex);
                }
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref createdRule);
                ComUtilities.Release(ref worksheetConditions!);
                ComUtilities.Release(ref cells!);
                ComUtilities.Release(ref formatConditions!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <summary>
    /// Adds a basic (cellValue/expression) rule and applies interior/font/border formatting.
    /// </summary>
    private static object AddBasicRule(
        Excel.FormatConditions formatConditions,
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
        Excel.FormatCondition? formatCondition = null;

        try
        {
            var xlType = normalizedType == "expression" ? 2 : 1;
            var xlOperator = ValidateAndParseBasicRuleArguments(
                normalizedType,
                operatorType,
                formula1,
                formula2);

            formatCondition = (Excel.FormatCondition)formatConditions.Add(
                Type: (Excel.XlFormatConditionType)xlType,
                Operator: xlOperator,
                Formula1: formula1 ?? "",
                Formula2: formula2 ?? "");

            ApplyRuleFormatting(formatCondition,
                interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
            object created = formatCondition;
            formatCondition = null;
            return created;
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
        object formatCondition,
        string? interiorColor,
        string? interiorPattern,
        string? fontColor,
        bool? fontBold,
        bool? fontItalic,
        string? borderStyle,
        string? borderColor)
    {
        Excel.Interior? interior = null;
        Excel.Font? font = null;
        Excel.Borders? borders = null;

        try
        {
            // Apply Interior formatting
            if (!string.IsNullOrEmpty(interiorColor) || !string.IsNullOrEmpty(interiorPattern))
            {
                interior = GetNativeInterior(formatCondition);
                if (!string.IsNullOrEmpty(interiorColor))
                    interior.Color = FormattingHelpers.ParseColor(interiorColor);
                if (!string.IsNullOrEmpty(interiorPattern))
                    interior.Pattern = ParseInteriorPattern(interiorPattern);
            }

            // Apply Font formatting
            if (!string.IsNullOrEmpty(fontColor) || fontBold.HasValue || fontItalic.HasValue)
            {
                font = GetNativeFont(formatCondition);
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
                borders = GetNativeBorders(formatCondition);
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
                    Excel.Border? border = null;
                    try
                    {
                        border = borders[(Excel.XlBordersIndex)borderIndex];
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
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.FormatConditions? formatConditions = null;

            try
            {
                // Get sheet
                sheet = RangeHelpers.ResolveWorksheet(ctx.Book, sheetName);

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
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.FormatConditions? formatConditions = null;

            try
            {
                sheet = RangeHelpers.ResolveWorksheet(ctx.Book, sheetName);

                range = sheet.Range[rangeAddress];
                formatConditions = range.FormatConditions;

                result.SheetName = sheet.Name;
                result.Rules = ReadFormatConditions(formatConditions, ct);
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
            Excel.Worksheet? sheet = null;
            Excel.Range? cells = null;
            Excel.FormatConditions? formatConditions = null;

            try
            {
                sheet = RangeHelpers.ResolveWorksheet(ctx.Book, sheetName);

                cells = sheet.Cells;
                formatConditions = cells.FormatConditions;

                result.SheetName = sheet.Name;
                result.Rules = ReadFormatConditions(formatConditions, ct);
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
    private static List<ConditionalFormatRuleInfo> ReadFormatConditions(Excel.FormatConditions formatConditions, CancellationToken cancellationToken)
    {
        var rules = new List<ConditionalFormatRuleInfo>();

        int count = Convert.ToInt32(formatConditions.Count, System.Globalization.CultureInfo.InvariantCulture);
        for (int i = 1; i <= count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            object? fc = null;
            Excel.Range? appliesTo = null;
            Excel.Interior? interior = null;
            Excel.Font? font = null;
            Excel.Borders? borders = null;
            Excel.Border? edgeBorder = null;

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
                    appliesTo = GetNativeAppliesTo(fc);
                    rule.AppliesTo = appliesTo?.Address;
                }
                catch (Exception ex) when (IsComOrBinderException(ex)) { }

                // Interior (fill)
                try
                {
                    interior = GetNativeInterior(fc);
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
                    font = GetNativeFont(fc);
                    object? bold = font.Bold;
                    object? italic = font.Italic;
                    rule.FontBold = bold is null ? null : Convert.ToBoolean(bold, System.Globalization.CultureInfo.InvariantCulture);
                    rule.FontItalic = italic is null ? null : Convert.ToBoolean(italic, System.Globalization.CultureInfo.InvariantCulture);
                    try
                    {
                        int fontColorIndex = Convert.ToInt32(font.ColorIndex, System.Globalization.CultureInfo.InvariantCulture);
                        if (fontColorIndex != -4105 && fontColorIndex != -4142) // not Automatic/None
                        {
                            rule.FontColor = FormattingHelpers.ColorToHex(Convert.ToInt32(font.Color, System.Globalization.CultureInfo.InvariantCulture));
                        }
                    }
                    catch (Exception ex) when (IsComOrBinderException(ex)) { }
                }
                catch (Exception ex) when (IsComOrBinderException(ex)) { }

                // Borders: scan all four edges and use the first that has a
                // style (rules typically apply borders uniformly, but external
                // rules may set only some edges).
                try
                {
                    borders = GetNativeBorders(fc);
                    // NOTE: FormatCondition.Borders is a 4-item collection indexed 1-4
                    // (left/top/bottom/right) - see write-side note in AddRule above.
                    foreach (int edgeIndex in new[] { 1, 2, 3, 4 }) // left, top, bottom, right
                    {
                        edgeBorder = borders[(Excel.XlBordersIndex)edgeIndex];
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
                switch (fc)
                {
                    case Excel.ColorScale item: rule.ColorScaleCriteria = ReadColorScale(item); break;
                    case Excel.Databar item: rule.DataBar = ReadDataBar(item); break;
                    case Excel.Top10 item: rule.Top10 = ReadTop10(item); break;
                    case Excel.IconSetCondition item: rule.IconSet = ReadIconSet(item); break;
                    case Excel.FormatCondition item when typeNum == 11: rule.DatePeriod = ReadTimePeriod(item); break;
                    case Excel.AboveAverage item: rule.AboveBelow = ReadAboveBelow(item); break;
                }

                if (fc is Excel.UniqueValues uniqueValues)
                    rule.DuplicateValues = Convert.ToInt32(uniqueValues.DupeUnique, System.Globalization.CultureInfo.InvariantCulture) == 1;
                if (fc is Excel.AboveAverage aboveAverage)
                    rule.StandardDeviations = Convert.ToInt32(aboveAverage.NumStdDev, System.Globalization.CultureInfo.InvariantCulture);
                rule.Fingerprint = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(
                    System.Text.Json.JsonSerializer.SerializeToUtf8Bytes(rule)));
                rules.Add(rule);
            }
            finally
            {
                ComUtilities.Release(ref edgeBorder!);
                ComUtilities.Release(ref borders!);
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref interior!);
                ComUtilities.Release(ref appliesTo!);
                ComUtilities.Release(ref fc!);
            }
        }

        return rules;
    }

    private static int ReadRuleTypeInt(object fc)
    {
        try
        {
            return fc switch
            {
                Excel.FormatCondition item => (int)item.Type,
                Excel.ColorScale item => (int)item.Type,
                Excel.Databar item => (int)item.Type,
                Excel.Top10 item => (int)item.Type,
                Excel.IconSetCondition item => (int)item.Type,
                Excel.UniqueValues item => (int)item.Type,
                Excel.AboveAverage item => (int)item.Type,
                _ => -1
            };
        }
        catch (Exception ex) when (IsComOrBinderException(ex)) { return -1; }
    }

    private static string? ReadRuleOperator(object fc)
    {
        try { return fc is Excel.FormatCondition item ? ConditionalFormattingOperatorToString(Convert.ToInt32(item.Operator, System.Globalization.CultureInfo.InvariantCulture)) : null; }
        catch (Exception ex) when (IsComOrBinderException(ex)) { return null; }
    }

    private static string? ReadRuleString(object fc, string property)
    {
        try
        {
            string? value = (fc, property) switch
            {
                (Excel.FormatCondition item, "Formula1") => item.Formula1,
                (Excel.FormatCondition item, "Formula2") => item.Formula2,
                _ => null
            };
            return string.IsNullOrEmpty(value) ? null : value;
        }
        catch (Exception ex) when (IsComOrBinderException(ex))
        {
            return null;
        }
    }

    private static int? ReadRuleInt(object fc, string property)
    {
        try
        {
            object? value = property switch
            {
                "Priority" when fc is Excel.FormatCondition or Excel.ColorScale or Excel.Databar
                    or Excel.IconSetCondition or Excel.Top10 or Excel.AboveAverage or Excel.UniqueValues
                    => GetNativePriority(fc),
                _ => null
            };
            return value == null ? null : (int?)Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (IsComOrBinderException(ex))
        {
            return null;
        }
    }

    private static bool? ReadRuleBool(object obj, string property)
    {
        try
        {
            object? value = (obj, property) switch
            {
                (Excel.FormatCondition item, "StopIfTrue") => item.StopIfTrue,
                (Excel.Top10 item, "StopIfTrue") => item.StopIfTrue,
                (Excel.AboveAverage item, "StopIfTrue") => item.StopIfTrue,
                (Excel.UniqueValues item, "StopIfTrue") => item.StopIfTrue,
                (Excel.ColorScale item, "StopIfTrue") => item.StopIfTrue,
                (Excel.Databar item, "StopIfTrue") => item.StopIfTrue,
                (Excel.IconSetCondition item, "StopIfTrue") => item.StopIfTrue,
                (Excel.Font item, "Bold") => item.Bold,
                (Excel.Font item, "Italic") => item.Italic,
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

    private static Excel.Range GetNativeAppliesTo(object rule) => rule switch
    {
        Excel.FormatCondition item => item.AppliesTo,
        Excel.ColorScale item => item.AppliesTo,
        Excel.Databar item => item.AppliesTo,
        Excel.IconSetCondition item => item.AppliesTo,
        Excel.Top10 item => item.AppliesTo,
        Excel.AboveAverage item => item.AppliesTo,
        Excel.UniqueValues item => item.AppliesTo,
        _ => throw new InvalidCastException("This native conditional-rule type has no supported AppliesTo range.")
    };

    private static Excel.Interior GetNativeInterior(object rule) => rule switch
    {
        Excel.FormatCondition item => item.Interior,
        Excel.Top10 item => item.Interior,
        Excel.AboveAverage item => item.Interior,
        Excel.UniqueValues item => item.Interior,
        _ => throw new InvalidCastException("This native conditional-rule type has no cell interior.")
    };

    private static Excel.Font GetNativeFont(object rule) => rule switch
    {
        Excel.FormatCondition item => item.Font,
        Excel.Top10 item => item.Font,
        Excel.AboveAverage item => item.Font,
        Excel.UniqueValues item => item.Font,
        _ => throw new InvalidCastException("This native conditional-rule type has no cell font.")
    };

    private static Excel.Borders GetNativeBorders(object rule) => rule switch
    {
        Excel.FormatCondition item => item.Borders,
        Excel.Top10 item => item.Borders,
        Excel.AboveAverage item => item.Borders,
        Excel.UniqueValues item => item.Borders,
        _ => throw new InvalidCastException("This native conditional-rule type has no cell borders.")
    };

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
