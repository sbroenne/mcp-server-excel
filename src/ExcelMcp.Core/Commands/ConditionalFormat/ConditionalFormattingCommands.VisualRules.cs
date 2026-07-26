using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Write-side helpers for creating visual conditional formatting rules
/// (colorScale, dataBar, iconSet, top10, aboveAverage, uniqueValues, timePeriod).
/// Each visual rule type requires a dedicated Excel COM Add* method rather than the
/// generic FormatConditions.Add used by cellValue/expression rules.
/// </summary>
public partial class ConditionalFormattingCommands
{
    /// <summary>
    /// Normalizes a rule-type string to a canonical lower-case key (kebab-case and
    /// camelCase both accepted, e.g. "color-scale" and "colorScale" -> "colorscale").
    /// </summary>
    private static string NormalizeRuleType(string type) =>
        (type ?? string.Empty).ToLowerInvariant().Replace("-", "");

    // === colorScale ===

    private static void AddColorScaleRule(
        dynamic formatConditions,
        string? minType, string? minValue, string? minColor,
        string? midType, string? midValue, string? midColor,
        string? maxType, string? maxValue, string? maxColor)
    {
        bool threeColor = !string.IsNullOrEmpty(midType)
            || !string.IsNullOrEmpty(midValue)
            || !string.IsNullOrEmpty(midColor);

        dynamic? colorScale = null;
        dynamic? criteria = null;
        try
        {
            colorScale = formatConditions.AddColorScale(ColorScaleType: threeColor ? 3 : 2);
            criteria = colorScale.ColorScaleCriteria;

            // Criterion 1 = minimum stop.
            SetColorScaleCriterion(criteria, 1, minType, minValue, minColor, defaultType: 1 /* lowest */);

            if (threeColor)
            {
                SetColorScaleCriterion(criteria, 2, midType, midValue, midColor,
                    defaultType: 4 /* percentile */, defaultValue: "50");
                SetColorScaleCriterion(criteria, 3, maxType, maxValue, maxColor, defaultType: 2 /* highest */);
            }
            else
            {
                SetColorScaleCriterion(criteria, 2, maxType, maxValue, maxColor, defaultType: 2 /* highest */);
            }
        }
        finally
        {
            ComUtilities.Release(ref criteria!);
            ComUtilities.Release(ref colorScale!);
        }
    }

    private static void SetColorScaleCriterion(
        dynamic criteria, int index,
        string? type, string? value, string? color,
        int defaultType, string? defaultValue = null)
    {
        dynamic? criterion = null;
        dynamic? formatColor = null;
        try
        {
            criterion = criteria.Item(index);

            int t = string.IsNullOrEmpty(type) ? defaultType : ParseConditionValueType(type);
            criterion.Type = t;

            string? v = value ?? (string.IsNullOrEmpty(type) ? defaultValue : null);
            if (ConditionValueTypeUsesValue(t) && !string.IsNullOrEmpty(v))
            {
                criterion.Value = ParseCriterionValue(v);
            }

            if (!string.IsNullOrEmpty(color))
            {
                formatColor = criterion.FormatColor;
                formatColor.Color = FormattingHelpers.ParseColor(color);
            }
        }
        finally
        {
            ComUtilities.Release(ref formatColor!);
            ComUtilities.Release(ref criterion!);
        }
    }

    // === dataBar ===

    private static void AddDataBarRule(
        dynamic formatConditions,
        string? fillColor, string? negativeColor, string? direction, bool? showValue,
        string? minType, string? minValue, string? maxType, string? maxValue)
    {
        dynamic? dataBar = null;
        try
        {
            dataBar = formatConditions.AddDatabar();

            if (!string.IsNullOrEmpty(fillColor))
            {
                dynamic? barColor = null;
                try
                {
                    barColor = dataBar.BarColor;
                    barColor.Color = FormattingHelpers.ParseColor(fillColor);
                }
                finally { ComUtilities.Release(ref barColor!); }
            }

            if (!string.IsNullOrEmpty(negativeColor))
            {
                dynamic? negativeFormat = null;
                dynamic? negColor = null;
                try
                {
                    negativeFormat = dataBar.NegativeBarFormat;
                    negativeFormat.ColorType = 0; // xlDataBarColor
                    negColor = negativeFormat.Color;
                    negColor.Color = FormattingHelpers.ParseColor(negativeColor);
                }
                finally
                {
                    ComUtilities.Release(ref negColor!);
                    ComUtilities.Release(ref negativeFormat!);
                }
            }

            if (!string.IsNullOrEmpty(direction))
                dataBar.Direction = ParseDataBarDirection(direction);

            if (showValue.HasValue)
                dataBar.ShowValue = showValue.Value;

            if (!string.IsNullOrEmpty(minType) || !string.IsNullOrEmpty(minValue))
                ModifyConditionValue(dataBar, "MinPoint", minType, minValue);

            if (!string.IsNullOrEmpty(maxType) || !string.IsNullOrEmpty(maxValue))
                ModifyConditionValue(dataBar, "MaxPoint", maxType, maxValue);
        }
        finally
        {
            ComUtilities.Release(ref dataBar!);
        }
    }

    private static void ModifyConditionValue(dynamic owner, string pointName, string? type, string? value)
    {
        dynamic? point = null;
        try
        {
            point = pointName == "MinPoint" ? owner.MinPoint : owner.MaxPoint;
            int t = string.IsNullOrEmpty(type)
                ? (pointName == "MinPoint" ? 1 /* lowest */ : 2 /* highest */)
                : ParseConditionValueType(type);

            if (ConditionValueTypeUsesValue(t) && !string.IsNullOrEmpty(value))
                point.Modify(t, ParseCriterionValue(value));
            else
                point.Modify(t);
        }
        finally
        {
            ComUtilities.Release(ref point!);
        }
    }

    // === iconSet ===

    private static void AddIconSetRule(
        dynamic workbook,
        dynamic formatConditions,
        string? iconSetId,
        bool? reverse,
        bool? showIconOnly,
        (string? Type, string? Value)[] thresholds)
    {
        dynamic? iconSetCondition = null;
        dynamic? iconSets = null;
        dynamic? iconSet = null;
        dynamic? iconCriteria = null;
        try
        {
            iconSetCondition = formatConditions.AddIconSetCondition();

            if (!string.IsNullOrEmpty(iconSetId))
            {
                iconSets = workbook.IconSets;
                iconSet = iconSets[ParseIconSetId(iconSetId)];
                iconSetCondition.IconSet = iconSet;
            }

            if (reverse.HasValue)
                iconSetCondition.ReverseOrder = reverse.Value;

            if (showIconOnly.HasValue)
                iconSetCondition.ShowIconOnly = showIconOnly.Value;

            iconCriteria = iconSetCondition.IconCriteria;
            int criteriaCount = Convert.ToInt32(iconCriteria.Count, CultureInfo.InvariantCulture);

            // IconCriteria item 1 is the implicit lower bound; editable thresholds start at item 2.
            for (int i = 0; i < thresholds.Length; i++)
            {
                var (tType, tValue) = thresholds[i];
                if (string.IsNullOrEmpty(tType) && string.IsNullOrEmpty(tValue))
                    continue;

                int itemIndex = i + 2;
                if (itemIndex > criteriaCount)
                    break;

                dynamic? criterion = null;
                try
                {
                    criterion = iconCriteria.Item(itemIndex);
                    if (!string.IsNullOrEmpty(tType))
                        criterion.Type = ParseConditionValueType(tType);
                    if (!string.IsNullOrEmpty(tValue))
                        criterion.Value = ParseCriterionValue(tValue);
                }
                finally
                {
                    ComUtilities.Release(ref criterion!);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref iconCriteria!);
            ComUtilities.Release(ref iconSet!);
            ComUtilities.Release(ref iconSets!);
            ComUtilities.Release(ref iconSetCondition!);
        }
    }

    // === top10 ===

    private static void AddTop10Rule(
        dynamic formatConditions,
        int? rank, bool? percent, string? topBottom,
        string? interiorColor, string? interiorPattern, string? fontColor,
        bool? fontBold, bool? fontItalic, string? borderStyle, string? borderColor)
    {
        dynamic? fc = null;
        try
        {
            fc = formatConditions.AddTop10();
            if (rank.HasValue)
                fc.Rank = rank.Value;
            if (percent.HasValue)
                fc.Percent = percent.Value;
            if (!string.IsNullOrEmpty(topBottom))
                fc.TopBottom = ParseTopBottom(topBottom);

            ApplyRuleFormatting(fc, interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
        }
        finally
        {
            ComUtilities.Release(ref fc!);
        }
    }

    // === aboveAverage ===

    private static void AddAboveAverageRule(
        dynamic formatConditions,
        string? aboveBelow,
        string? interiorColor, string? interiorPattern, string? fontColor,
        bool? fontBold, bool? fontItalic, string? borderStyle, string? borderColor)
    {
        dynamic? fc = null;
        try
        {
            fc = formatConditions.AddAboveAverage();
            if (!string.IsNullOrEmpty(aboveBelow))
                fc.AboveBelow = ParseAboveBelow(aboveBelow);

            ApplyRuleFormatting(fc, interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
        }
        finally
        {
            ComUtilities.Release(ref fc!);
        }
    }

    // === uniqueValues ===

    private static void AddUniqueValuesRule(
        dynamic formatConditions,
        bool duplicate,
        string? interiorColor, string? interiorPattern, string? fontColor,
        bool? fontBold, bool? fontItalic, string? borderStyle, string? borderColor)
    {
        dynamic? fc = null;
        try
        {
            fc = formatConditions.AddUniqueValues();
            fc.DupeUnique = duplicate ? 1 /* xlDuplicate */ : 0 /* xlUnique */;

            ApplyRuleFormatting(fc, interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
        }
        finally
        {
            ComUtilities.Release(ref fc!);
        }
    }

    // === timePeriod ===

    private static void AddTimePeriodRule(
        dynamic formatConditions,
        string? datePeriod,
        string? interiorColor, string? interiorPattern, string? fontColor,
        bool? fontBold, bool? fontItalic, string? borderStyle, string? borderColor)
    {
        if (string.IsNullOrEmpty(datePeriod))
            throw new ArgumentException("datePeriod is required for timePeriod rules (e.g. today, last7Days, thisMonth).");

        dynamic? fc = null;
        try
        {
            fc = formatConditions.Add(Type: 11 /* xlTimePeriod */, DateOperator: ParseTimePeriod(datePeriod));
            ApplyRuleFormatting(fc, interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
        }
        finally
        {
            ComUtilities.Release(ref fc!);
        }
    }

    // === blanksCondition and other simple type-only rules ===

    private static void AddSimpleRule(
        dynamic formatConditions,
        int xlType,
        string? interiorColor, string? interiorPattern, string? fontColor,
        bool? fontBold, bool? fontItalic, string? borderStyle, string? borderColor)
    {
        dynamic? fc = null;
        try
        {
            fc = formatConditions.Add(Type: xlType);
            ApplyRuleFormatting(fc, interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
        }
        finally
        {
            ComUtilities.Release(ref fc!);
        }
    }

    // === shared value helpers ===

    /// <summary>
    /// Whether a threshold type carries an associated value. Minimum/maximum/automatic
    /// stops derive their value from the data, so no explicit value is set.
    /// </summary>
    private static bool ConditionValueTypeUsesValue(int conditionValueType) =>
        conditionValueType is 0 /* number */ or 3 /* percent */ or 4 /* percentile */ or 5 /* formula */;

    /// <summary>
    /// Parses a threshold value: numeric strings become doubles; anything else (e.g. a
    /// formula) is passed through as a string.
    /// </summary>
    private static object ParseCriterionValue(string value) =>
        double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)
            ? d
            : value;

    // === parse helpers (string -> Excel COM enum int) ===

    private static int ParseConditionValueType(string type)
    {
        return type.ToLowerInvariant().Replace("-", "").Replace(" ", "") switch
        {
            "none" => -1,          // xlConditionValueNone
            "number" or "num" => 0, // xlConditionValueNumber
            "lowest" or "lowestvalue" or "minimum" or "min" => 1, // xlConditionValueLowestValue
            "highest" or "highestvalue" or "maximum" or "max" => 2, // xlConditionValueHighestValue
            "percent" => 3,        // xlConditionValuePercent
            "percentile" => 4,     // xlConditionValuePercentile
            "formula" => 5,        // xlConditionValueFormula
            "automaticmin" or "automaticminimum" => 6, // xlConditionValueAutomaticMin
            "automaticmax" or "automaticmaximum" => 7, // xlConditionValueAutomaticMax
            _ => throw new ArgumentException(
                $"Invalid threshold type: '{type}'. Valid values: number, minimum, maximum, percent, percentile, formula, automaticMinimum, automaticMaximum")
        };
    }

    private static int ParseDataBarDirection(string direction)
    {
        return direction.ToLowerInvariant().Replace("-", "").Replace(" ", "") switch
        {
            "context" => -5002,      // xlContext (XlReadingOrder)
            "lefttoright" or "ltr" => -5003, // xlLTR
            "righttoleft" or "rtl" => -5004, // xlRTL
            _ => throw new ArgumentException(
                $"Invalid data bar direction: '{direction}'. Valid values: context, leftToRight, rightToLeft")
        };
    }

    private static int ParseIconSetId(string iconSetId)
    {
        return iconSetId.ToLowerInvariant().Replace("-", "").Replace(" ", "") switch
        {
            "3arrows" => 1,          // xl3Arrows
            "3arrowsgray" => 2,      // xl3ArrowsGray
            "3flags" => 3,           // xl3Flags
            "3trafficlights1" => 4,  // xl3TrafficLights1
            "3trafficlights2" => 5,  // xl3TrafficLights2
            "3signs" => 6,           // xl3Signs
            "3symbols" => 7,         // xl3Symbols
            "4arrows" => 8,          // xl4Arrows
            "4arrowsgray" => 9,      // xl4ArrowsGray
            "4redtoblack" => 10,     // xl4RedToBlack
            "4crv" or "4ratings" => 11, // xl4CRV (Excel UI: "4 Ratings")
            "4trafficlights" => 12,  // xl4TrafficLights
            "5arrows" => 13,         // xl5Arrows
            "5arrowsgray" => 14,     // xl5ArrowsGray
            "5crv" or "5ratings" => 15, // xl5CRV (Excel UI: "5 Ratings")
            "5quarters" => 16,       // xl5Quarters
            _ => throw new ArgumentException(
                $"Invalid icon set id: '{iconSetId}'. Valid values include: 3Arrows, 3ArrowsGray, 3Flags, 3TrafficLights1, 3TrafficLights2, 3Signs, 3Symbols, 4Arrows, 4ArrowsGray, 4RedToBlack, 4Ratings, 4TrafficLights, 5Arrows, 5ArrowsGray, 5Ratings, 5Quarters")
        };
    }

    private static int ParseTopBottom(string topBottom)
    {
        return topBottom.ToLowerInvariant() switch
        {
            "bottom" => 0, // xlTop10Bottom
            "top" => 1,    // xlTop10Top
            _ => throw new ArgumentException($"Invalid topBottom value: '{topBottom}'. Valid values: top, bottom")
        };
    }

    private static int ParseAboveBelow(string aboveBelow)
    {
        return aboveBelow.ToLowerInvariant().Replace("-", "").Replace(" ", "") switch
        {
            "aboveaverage" or "above" => 0,  // xlAboveAverage
            "abovestddev" => 1,              // xlAboveStdDev
            "belowaverage" or "below" => 2,  // xlBelowAverage
            "belowstddev" => 3,              // xlBelowStdDev
            "equalaboveaverage" => 4,        // xlEqualAboveAverage
            "equalbelowaverage" => 5,        // xlEqualBelowAverage
            _ => throw new ArgumentException(
                $"Invalid aboveBelow value: '{aboveBelow}'. Valid values: aboveAverage, belowAverage, aboveStdDev, belowStdDev, equalAboveAverage, equalBelowAverage")
        };
    }

    private static int ParseTimePeriod(string datePeriod)
    {
        return datePeriod.ToLowerInvariant().Replace("-", "").Replace(" ", "") switch
        {
            "today" => 0,      // xlToday
            "yesterday" => 1,  // xlYesterday
            "last7days" => 2,  // xlLast7Days
            "thisweek" => 3,   // xlThisWeek
            "lastweek" => 4,   // xlLastWeek
            "nextweek" => 5,   // xlNextWeek
            "thismonth" => 6,  // xlThisMonth
            "lastmonth" => 7,  // xlLastMonth
            "nextmonth" => 8,  // xlNextMonth
            "tomorrow" => 9,   // xlTomorrow
            _ => throw new ArgumentException(
                $"Invalid datePeriod value: '{datePeriod}'. Valid values: today, yesterday, tomorrow, last7Days, thisWeek, lastWeek, nextWeek, thisMonth, lastMonth, nextMonth")
        };
    }
}
