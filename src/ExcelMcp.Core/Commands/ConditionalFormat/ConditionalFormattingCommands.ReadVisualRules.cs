using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Read-side helpers that extract type-specific configuration from visual conditional
/// formatting rules (colorScale, dataBar, iconSet, top10, aboveAverage, timePeriod).
/// Known unavailable members degrade to null; unexpected native failures propagate.
/// </summary>
public partial class ConditionalFormattingCommands
{
    private static List<ColorScaleCriterionInfo>? ReadColorScale(Excel.ColorScale fc)
    {
        Excel.ColorScaleCriteria? criteria = null;
        try
        {
            criteria = fc.ColorScaleCriteria;
            int count = Convert.ToInt32(criteria.Count, CultureInfo.InvariantCulture);
            var list = new List<ColorScaleCriterionInfo>();

            for (int i = 1; i <= count; i++)
            {
                Excel.ColorScaleCriterion? criterion = null;
                Excel.FormatColor? formatColor = null;
                try
                {
                    criterion = criteria[i];
                    var info = new ColorScaleCriterionInfo();

                    int t = (int)criterion.Type;
                    info.Type = ConditionValueTypeToString(t);

                    if (ConditionValueTypeUsesValue(t))
                        info.Value = ReadConditionValue(() => criterion.Value);

                    try
                    {
                        formatColor = criterion.FormatColor;
                        info.Color = FormattingHelpers.ColorToHex(Convert.ToInt32((object)formatColor.Color, CultureInfo.InvariantCulture));
                    }
                    catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }

                    list.Add(info);
                }
                finally
                {
                    ComUtilities.Release(ref formatColor!);
                    ComUtilities.Release(ref criterion!);
                }
            }

            return list.Count > 0 ? list : null;
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { return null; }
        finally
        {
            ComUtilities.Release(ref criteria!);
        }
    }

    private static DataBarInfo? ReadDataBar(Excel.Databar fc)
    {
        try
        {
            var info = new DataBarInfo();

            Excel.FormatColor? barColor = null;
            try
            {
                barColor = fc.BarColor;
                info.FillColor = FormattingHelpers.ColorToHex(Convert.ToInt32((object)barColor.Color, CultureInfo.InvariantCulture));
            }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }
            finally { ComUtilities.Release(ref barColor!); }

            // Negative bar color only meaningful when a custom color type is used.
            Excel.NegativeBarFormat? negativeFormat = null;
            Excel.FormatColor? negColor = null;
            try
            {
                negativeFormat = fc.NegativeBarFormat;
                int colorType = (int)negativeFormat.ColorType;
                if (colorType == 0) // xlDataBarColor (explicit color, not "same as positive")
                {
                    negColor = negativeFormat.Color;
                    info.BarColorNegative = FormattingHelpers.ColorToHex(Convert.ToInt32((object)negColor.Color, CultureInfo.InvariantCulture));
                }
            }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }
            finally
            {
                ComUtilities.Release(ref negColor!);
                ComUtilities.Release(ref negativeFormat!);
            }

            try { info.Direction = DataBarDirectionToString((int)fc.Direction); }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }

            try { info.ShowValue = Convert.ToBoolean(fc.ShowValue, CultureInfo.InvariantCulture); }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }

            ReadDataBarPoint(fc, info, isMin: true);
            ReadDataBarPoint(fc, info, isMin: false);

            return info;
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { return null; }
    }

    private static void ReadDataBarPoint(Excel.Databar fc, DataBarInfo info, bool isMin)
    {
        Excel.ConditionValue? point = null;
        try
        {
            point = isMin ? fc.MinPoint : fc.MaxPoint;
            int t = (int)point.Type;
            string typeStr = ConditionValueTypeToString(t);
            string? valueStr = ConditionValueTypeUsesValue(t) ? ReadConditionValue(() => point.Value) : null;

            if (isMin) { info.MinType = typeStr; info.MinValue = valueStr; }
            else { info.MaxType = typeStr; info.MaxValue = valueStr; }
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }
        finally
        {
            ComUtilities.Release(ref point!);
        }
    }

    private static IconSetInfo? ReadIconSet(Excel.IconSetCondition fc)
    {
        try
        {
            var info = new IconSetInfo();

            Excel.IconSet? iconSet = null;
            try
            {
                iconSet = (Excel.IconSet)fc.IconSet;
                info.Id = IconSetIdToString((int)iconSet.ID);
            }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }
            finally { ComUtilities.Release(ref iconSet!); }

            try { info.Reverse = Convert.ToBoolean(fc.ReverseOrder, CultureInfo.InvariantCulture); }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }

            try { info.ShowIconOnly = Convert.ToBoolean(fc.ShowIconOnly, CultureInfo.InvariantCulture); }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }

            Excel.IconCriteria? iconCriteria = null;
            try
            {
                iconCriteria = fc.IconCriteria;
                int count = Convert.ToInt32(iconCriteria.Count, CultureInfo.InvariantCulture);
                var criteria = new List<IconCriterionInfo>();

                for (int i = 1; i <= count; i++)
                {
                    Excel.IconCriterion? criterion = null;
                    try
                    {
                        criterion = iconCriteria[i];
                        var c = new IconCriterionInfo();

                        // The first icon has no threshold; Excel rejects its threshold properties.
                        if (i > 1)
                        {
                            c.Operator = ConditionalFormattingOperatorToString(Convert.ToInt32(criterion.Operator, CultureInfo.InvariantCulture));
                            int t = (int)criterion.Type;
                            c.Type = ConditionValueTypeToString(t);
                            if (ConditionValueTypeUsesValue(t))
                                c.Value = ReadConditionValue(() => criterion.Value);
                        }

                        try { c.Icon = IconToString((int)criterion.Icon); }
                        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }

                        criteria.Add(c);
                    }
                    finally
                    {
                        ComUtilities.Release(ref criterion!);
                    }
                }

                if (criteria.Count > 0)
                    info.Criteria = criteria;
            }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }
            finally { ComUtilities.Release(ref iconCriteria!); }

            return info;
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { return null; }
    }

    private static Top10Info? ReadTop10(Excel.Top10 fc)
    {
        try
        {
            var info = new Top10Info();

            try { info.Rank = Convert.ToInt32(fc.Rank, CultureInfo.InvariantCulture); }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }

            try { info.Percent = Convert.ToBoolean(fc.Percent, CultureInfo.InvariantCulture); }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }

            try { info.TopBottom = TopBottomToString((int)fc.TopBottom); }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { }

            return info;
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { return null; }
    }

    private static string? ReadAboveBelow(Excel.AboveAverage fc)
    {
        try { return AboveBelowToString((int)fc.AboveBelow); }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { return null; }
    }

    private static string? ReadTimePeriod(Excel.FormatCondition fc)
    {
        try { return TimePeriodToString(Convert.ToInt32(fc.DateOperator, CultureInfo.InvariantCulture)); }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { return null; }
    }

    /// <summary>
    /// Reads a ConditionValue/criterion <c>Value</c> and formats it as an invariant string.
    /// Integral doubles are rendered without a trailing ".0". Returns null when unreadable.
    /// </summary>
    private static string? ReadConditionValue(Func<object?> readValue)
    {
        try
        {
            object? value = readValue();
            if (value == null)
                return null;

            if (value is double d)
            {
                return d == Math.Floor(d) && !double.IsInfinity(d)
                    ? ((long)d).ToString(CultureInfo.InvariantCulture)
                    : d.ToString(CultureInfo.InvariantCulture);
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex)) { return null; }
    }

    // === reverse mappings (int -> string) ===

    private static string ConditionValueTypeToString(int type)
    {
        return type switch
        {
            -1 => "none",              // xlConditionValueNone
            0 => "number",             // xlConditionValueNumber
            1 => "minimum",            // xlConditionValueLowestValue
            2 => "maximum",            // xlConditionValueHighestValue
            3 => "percent",            // xlConditionValuePercent
            4 => "percentile",         // xlConditionValuePercentile
            5 => "formula",            // xlConditionValueFormula
            6 => "automaticMinimum",   // xlConditionValueAutomaticMin
            7 => "automaticMaximum",   // xlConditionValueAutomaticMax
            _ => $"unknown({type})"
        };
    }

    private static string DataBarDirectionToString(int direction)
    {
        return direction switch
        {
            -5002 => "context",       // xlContext (XlReadingOrder)
            -5003 => "leftToRight",   // xlLTR
            -5004 => "rightToLeft",   // xlRTL
            _ => $"unknown({direction})"
        };
    }

    private static string TopBottomToString(int topBottom)
    {
        return topBottom switch
        {
            0 => "bottom", // xlTop10Bottom
            1 => "top",    // xlTop10Top
            _ => $"unknown({topBottom})"
        };
    }

    private static string AboveBelowToString(int aboveBelow)
    {
        return aboveBelow switch
        {
            0 => "aboveAverage",       // xlAboveAverage
            1 => "belowAverage",       // xlBelowAverage
            2 => "equalAboveAverage",  // xlEqualAboveAverage
            3 => "equalBelowAverage",  // xlEqualBelowAverage
            4 => "aboveStdDev",        // xlAboveStdDev
            5 => "belowStdDev",        // xlBelowStdDev
            _ => $"unknown({aboveBelow})"
        };
    }

    private static string TimePeriodToString(int period)
    {
        return period switch
        {
            0 => "today",     // xlToday
            1 => "yesterday", // xlYesterday
            2 => "last7Days", // xlLast7Days
            3 => "thisWeek",  // xlThisWeek
            4 => "lastWeek",  // xlLastWeek
            5 => "nextWeek",  // xlNextWeek
            6 => "thisMonth", // xlThisMonth
            7 => "lastMonth", // xlLastMonth
            8 => "nextMonth", // xlNextMonth
            9 => "tomorrow",  // xlTomorrow
            _ => $"unknown({period})"
        };
    }

    private static string IconSetIdToString(int id)
    {
        return id switch
        {
            1 => "3Arrows",          // xl3Arrows
            2 => "3ArrowsGray",      // xl3ArrowsGray
            3 => "3Flags",           // xl3Flags
            4 => "3TrafficLights1",  // xl3TrafficLights1
            5 => "3TrafficLights2",  // xl3TrafficLights2
            6 => "3Signs",           // xl3Signs
            7 => "3Symbols",         // xl3Symbols
            8 => "4Arrows",          // xl4Arrows
            9 => "4ArrowsGray",      // xl4ArrowsGray
            10 => "4RedToBlack",     // xl4RedToBlack
            11 => "4Ratings",        // xl4CRV
            12 => "4TrafficLights",  // xl4TrafficLights
            13 => "5Arrows",         // xl5Arrows
            14 => "5ArrowsGray",     // xl5ArrowsGray
            15 => "5Ratings",        // xl5CRV
            16 => "5Quarters",       // xl5Quarters
            _ => $"iconSet({id})"
        };
    }

    private static string IconToString(int icon)
    {
        return icon switch
        {
            -1 => "noCellIcon",                    // xlIconNoCellIcon
            1 => "greenUpArrow",                   // xlIconGreenUpArrow
            2 => "yellowSideArrow",                // xlIconYellowSideArrow
            3 => "redDownArrow",                   // xlIconRedDownArrow
            4 => "grayUpArrow",                    // xlIconGrayUpArrow
            5 => "graySideArrow",                  // xlIconGraySideArrow
            6 => "grayDownArrow",                  // xlIconGrayDownArrow
            7 => "greenFlag",                      // xlIconGreenFlag
            8 => "yellowFlag",                     // xlIconYellowFlag
            9 => "redFlag",                        // xlIconRedFlag
            10 => "greenCircle",                   // xlIconGreenCircle
            11 => "yellowCircle",                  // xlIconYellowCircle
            12 => "redCircleWithBorder",           // xlIconRedCircleWithBorder
            13 => "blackCircleWithBorder",         // xlIconBlackCircleWithBorder
            14 => "greenTrafficLight",             // xlIconGreenTrafficLight
            15 => "yellowTrafficLight",            // xlIconYellowTrafficLight
            16 => "redTrafficLight",               // xlIconRedTrafficLight
            17 => "yellowTriangle",                // xlIconYellowTriangle
            18 => "redDiamond",                    // xlIconRedDiamond
            19 => "greenCheckSymbol",              // xlIconGreenCheckSymbol
            20 => "yellowExclamationSymbol",       // xlIconYellowExclamationSymbol
            21 => "redCrossSymbol",                // xlIconRedCrossSymbol
            22 => "greenCheck",                    // xlIconGreenCheck
            23 => "yellowExclamation",             // xlIconYellowExclamation
            24 => "redCross",                      // xlIconRedCross
            25 => "yellowUpInclineArrow",          // xlIconYellowUpInclineArrow
            26 => "yellowDownInclineArrow",        // xlIconYellowDownInclineArrow
            27 => "grayUpInclineArrow",            // xlIconGrayUpInclineArrow
            28 => "grayDownInclineArrow",          // xlIconGrayDownInclineArrow
            29 => "redCircle",                     // xlIconRedCircle
            30 => "pinkCircle",                    // xlIconPinkCircle
            31 => "grayCircle",                    // xlIconGrayCircle
            32 => "blackCircle",                   // xlIconBlackCircle
            33 => "circleWithOneWhiteQuarter",     // xlIconCircleWithOneWhiteQuarter
            34 => "circleWithTwoWhiteQuarters",    // xlIconCircleWithTwoWhiteQuarters
            35 => "circleWithThreeWhiteQuarters",  // xlIconCircleWithThreeWhiteQuarters
            36 => "whiteCircleAllWhiteQuarters",   // xlIconWhiteCircleAllWhiteQuarters
            37 => "zeroBars",                      // xlIcon0Bars
            38 => "oneBar",                        // xlIcon1Bar
            39 => "twoBars",                       // xlIcon2Bars
            40 => "threeBars",                     // xlIcon3Bars
            41 => "fourBars",                      // xlIcon4Bars
            42 => "goldStar",                      // xlIconGoldStar
            43 => "halfGoldStar",                  // xlIconHalfGoldStar
            44 => "silverStar",                    // xlIconSilverStar
            45 => "greenUpTriangle",               // xlIconGreenUpTriangle
            46 => "yellowDash",                    // xlIconYellowDash
            47 => "redDownTriangle",               // xlIconRedDownTriangle
            48 => "fourFilledBoxes",               // xlIcon4FilledBoxes
            49 => "threeFilledBoxes",              // xlIcon3FilledBoxes
            50 => "twoFilledBoxes",                // xlIcon2FilledBoxes
            51 => "oneFilledBox",                  // xlIcon1FilledBox
            52 => "zeroFilledBoxes",               // xlIcon0FilledBoxes
            _ => $"icon({icon})"
        };
    }
}
