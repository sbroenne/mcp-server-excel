using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class ConditionalFormattingCommands
{
    private static void ValidateVisualUpdate(ConditionalFormatRuleInfo current, ConditionalRuleUpdateOptions options)
    {
        if (current.Type == "colorScale")
        {
            var stops = current.ColorScaleCriteria ??
                throw new InvalidOperationException("Excel did not expose the existing color-scale stops.");
            if (stops.Count is not (2 or 3))
                throw new NotSupportedException("The existing color scale does not have two or three stops.");
            ValidateThreshold(options.ColorScaleMinType, options.ColorScaleMinValue,
                stops[0].Type, stops[0].Value, [0, 1, 3, 4, 5]);
            ValidateThreshold(options.ColorScaleMaxType, options.ColorScaleMaxValue,
                stops[^1].Type, stops[^1].Value, [0, 2, 3, 4, 5]);
            if (options.ColorScaleMidType is not null || options.ColorScaleMidValue is not null || options.ColorScaleMidColor is not null)
            {
                if (stops.Count != 3)
                    throw new ArgumentException("A two-color scale has no midpoint; changing the stop count requires a new rule.", nameof(options));
                ValidateThreshold(options.ColorScaleMidType, options.ColorScaleMidValue,
                    stops[1].Type, stops[1].Value, [0, 3, 4, 5]);
            }
        }
        else if (current.Type == "dataBar")
        {
            var bar = current.DataBar ??
                throw new InvalidOperationException("Excel did not expose the existing data-bar settings.");
            ValidateThreshold(options.DataBarMinType, options.DataBarMinValue, bar.MinType, bar.MinValue, [0, 1, 3, 4, 5, 6]);
            ValidateThreshold(options.DataBarMaxType, options.DataBarMaxValue, bar.MaxType, bar.MaxValue, [0, 2, 3, 4, 5, 7]);
        }
        else if (current.Type == "iconSet")
        {
            var icons = current.IconSet ??
                throw new InvalidOperationException("Excel did not expose the existing icon-set settings.");
            var existing = icons.Criteria ??
                throw new InvalidOperationException("Excel did not expose the existing icon thresholds.");
            int count = options.IconSetId is null ? existing.Count : GetIconCount(ParseIconSetId(options.IconSetId));
            var thresholds = GetUpdatedIconThresholds(options);
            for (int i = 0; i < thresholds.Length; i++)
            {
                var (type, value) = thresholds[i];
                if (type is null && value is null)
                    continue;
                if (i + 1 >= count)
                    throw new ArgumentException("The requested icon threshold does not exist in the selected icon set.", nameof(options));
                string? currentType = i + 1 < existing.Count && options.IconSetId is null
                    ? existing[i + 1].Type : "percent";
                string? currentValue = i + 1 < existing.Count && options.IconSetId is null
                    ? existing[i + 1].Value : ((i + 1) * 100d / count).ToString(CultureInfo.InvariantCulture);
                ValidateThreshold(type, value, currentType, currentValue, [0, 3, 4, 5]);
            }
        }
    }

    private static void ValidateThreshold(string? type, string? value, string? currentType,
        string? currentValue, int[] allowedTypes)
    {
        if (type is null && value is null)
            return;
        int kind = ParseConditionValueType(type ?? currentType ??
            throw new InvalidOperationException("Excel did not expose the existing threshold type."));
        if (!allowedTypes.Contains(kind))
            throw new ArgumentException("The threshold type is not applicable to this stop.");
        if (!ConditionValueTypeUsesValue(kind))
        {
            if (value is not null)
                throw new ArgumentException("This threshold type does not accept an explicit value.");
            return;
        }
        string text = value ?? currentValue ??
            throw new ArgumentException("This threshold type requires an explicit value.");
        if (kind == 5)
        {
            if (!text.StartsWith('='))
                throw new ArgumentException("Formula thresholds must begin with '='.");
        }
        else
        {
            if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) ||
                !double.IsFinite(number))
                throw new ArgumentException("Numeric thresholds require a finite invariant number.");
            if (kind is 3 or 4 && number is < 0 or > 100)
                throw new ArgumentException("Percent and percentile thresholds must be between 0 and 100.");
        }
    }

    private static int GetIconCount(int id) => id <= 7 ? 3 : id <= 12 ? 4 : 5;

    private static (string? Type, string? Value)[] GetUpdatedIconThresholds(ConditionalRuleUpdateOptions options) =>
    [
        (options.IconThreshold1Type, options.IconThreshold1Value),
        (options.IconThreshold2Type, options.IconThreshold2Value),
        (options.IconThreshold3Type, options.IconThreshold3Value),
        (options.IconThreshold4Type, options.IconThreshold4Value)
    ];

    private static void UpdateColorScale(Excel.ColorScale scale, ConditionalRuleUpdateOptions options)
    {
        Excel.ColorScaleCriteria? criteria = null;
        try
        {
            criteria = scale.ColorScaleCriteria;
            UpdateColorStop(criteria, 1, options.ColorScaleMinType, options.ColorScaleMinValue, options.ColorScaleMinColor);
            if (criteria.Count == 3)
                UpdateColorStop(criteria, 2, options.ColorScaleMidType, options.ColorScaleMidValue, options.ColorScaleMidColor);
            UpdateColorStop(criteria, criteria.Count, options.ColorScaleMaxType, options.ColorScaleMaxValue, options.ColorScaleMaxColor);
        }
        finally { ComUtilities.Release(ref criteria); }
    }

    private static void UpdateColorStop(Excel.ColorScaleCriteria criteria, int index, string? type, string? value, string? color)
    {
        if (type is null && value is null && color is null)
            return;
        Excel.ColorScaleCriterion? stop = null;
        Excel.FormatColor? format = null;
        try
        {
            stop = criteria.Item[index];
            if (type is not null) stop.Type = (Excel.XlConditionValueTypes)ParseConditionValueType(type);
            if (value is not null) stop.Value = ParseCriterionValue(value);
            if (color is not null)
            {
                format = stop.FormatColor;
                format.Color = FormattingHelpers.ParseColor(color);
            }
        }
        finally
        {
            ComUtilities.Release(ref format);
            ComUtilities.Release(ref stop);
        }
    }

    private static void UpdateDataBar(Excel.Databar bar, ConditionalRuleUpdateOptions options)
    {
        Excel.FormatColor? fill = null;
        Excel.NegativeBarFormat? negative = null;
        Excel.FormatColor? negativeColor = null;
        Excel.ConditionValue? minimum = null;
        Excel.ConditionValue? maximum = null;
        try
        {
            if (options.DataBarColor is not null)
            {
                fill = bar.BarColor;
                fill.Color = FormattingHelpers.ParseColor(options.DataBarColor);
            }
            if (options.DataBarNegativeColor is not null)
            {
                negative = bar.NegativeBarFormat;
                negative.ColorType = Excel.XlDataBarNegativeColorType.xlDataBarColor;
                negativeColor = negative.Color;
                negativeColor.Color = FormattingHelpers.ParseColor(options.DataBarNegativeColor);
            }
            if (options.DataBarDirection is not null)
                bar.Direction = ParseDataBarDirection(options.DataBarDirection);
            if (options.DataBarShowValue.HasValue) bar.ShowValue = options.DataBarShowValue.Value;
            if (options.DataBarMinType is not null || options.DataBarMinValue is not null)
            {
                minimum = bar.MinPoint;
                UpdateConditionPoint(minimum, options.DataBarMinType, options.DataBarMinValue);
            }
            if (options.DataBarMaxType is not null || options.DataBarMaxValue is not null)
            {
                maximum = bar.MaxPoint;
                UpdateConditionPoint(maximum, options.DataBarMaxType, options.DataBarMaxValue);
            }
        }
        finally
        {
            ComUtilities.Release(ref maximum);
            ComUtilities.Release(ref minimum);
            ComUtilities.Release(ref negativeColor);
            ComUtilities.Release(ref negative);
            ComUtilities.Release(ref fill);
        }
    }

    private static void UpdateConditionPoint(Excel.ConditionValue point, string? type, string? value)
    {
        var kind = type is null ? point.Type : (Excel.XlConditionValueTypes)ParseConditionValueType(type);
        object? nativeValue = value is null && ConditionValueTypeUsesValue((int)kind) ? point.Value :
            value is null ? Type.Missing : ParseCriterionValue(value);
        point.Modify(kind, nativeValue);
    }

    private static void UpdateIconSet(Excel.IconSetCondition icons, Excel.Workbook book, ConditionalRuleUpdateOptions options)
    {
        Excel.IconSets? sets = null;
        Excel.IconSet? set = null;
        Excel.IconCriteria? criteria = null;
        try
        {
            if (options.IconSetId is not null)
            {
                sets = book.IconSets;
                set = sets.Item[ParseIconSetId(options.IconSetId)];
                icons.IconSet = set;
            }
            if (options.IconSetReverse.HasValue) icons.ReverseOrder = options.IconSetReverse.Value;
            if (options.IconSetShowIconOnly.HasValue) icons.ShowIconOnly = options.IconSetShowIconOnly.Value;
            criteria = icons.IconCriteria;
            var thresholds = GetUpdatedIconThresholds(options);
            for (int i = 0; i < thresholds.Length; i++)
            {
                var (type, value) = thresholds[i];
                if (type is null && value is null) continue;
                Excel.IconCriterion? criterion = null;
                try
                {
                    criterion = criteria.Item[i + 2];
                    if (type is not null) criterion.Type = (Excel.XlConditionValueTypes)ParseConditionValueType(type);
                    if (value is not null) criterion.Value = ParseCriterionValue(value);
                }
                finally { ComUtilities.Release(ref criterion); }
            }
        }
        finally
        {
            ComUtilities.Release(ref criteria);
            ComUtilities.Release(ref set);
            ComUtilities.Release(ref sets);
        }
    }
}
