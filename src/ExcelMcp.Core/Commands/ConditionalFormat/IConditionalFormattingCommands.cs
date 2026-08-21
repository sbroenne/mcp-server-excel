using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Conditional formatting - visual rules based on cell values.
/// TYPES: cellValue (requires operatorType+formula1), expression (formula only),
/// colorScale, dataBar, iconSet, top10, aboveAverage, timePeriod, uniqueValues, blanksCondition.
/// Both camelCase and kebab-case accepted.
/// FORMAT: interiorColor/fontColor as #RRGGBB, fontBold/Italic, borderStyle/Color.
/// Visual rule types use dedicated parameters on add-rule and return their type-specific
/// configuration from list-rules / list-worksheet-rules.
///
/// OPERATORS: equal, notEqual, greater, less, greaterEqual, lessEqual, between, notBetween.
/// For 'between' and 'notBetween', both formula1 and formula2 are required.
/// </summary>
[ServiceCategory("conditionalformat", "ConditionalFormat")]
[McpTool("conditionalformat", Title = "Conditional Formatting", Destructive = true, Category = "structure",
    Description = "Conditional formatting - visual rules based on cell values. TYPES: cellValue, expression, colorScale, dataBar, iconSet, top10, aboveAverage, timePeriod, uniqueValues, blanksCondition (accepts both camelCase and kebab-case). For cellValue: requires operatorType + formula1. Visual types use dedicated add-rule parameters and list-rules returns their type-specific config (colorScaleCriteria, dataBar, iconSet, top10, aboveBelow, datePeriod). FORMAT: interiorColor/fontColor as #RRGGBB hex, fontBold/fontItalic booleans, borderStyle/borderColor.")]
public interface IConditionalFormattingCommands
{
    /// <summary>
    /// Adds a conditional formatting rule to a range.
    /// Excel COM: Range.FormatConditions.Add / AddColorScale / AddDatabar /
    /// AddIconSetCondition / AddTop10 / AddAboveAverage / AddUniqueValues.
    ///
    /// Basic rules (cellValue, expression) use operatorType/formula1/formula2 plus the
    /// interior/font/border format parameters. Visual rules use their own dedicated
    /// parameters (all optional; only the ones matching ruleType are read):
    /// - colorScale: colorScaleMin/Mid/Max Type/Value/Color (supply a mid stop for a 3-color scale)
    /// - dataBar: dataBarColor, dataBarNegativeColor, dataBarDirection, dataBarShowValue,
    ///   dataBarMin/Max Type/Value
    /// - iconSet: iconSetId, iconSetReverse, iconSetShowIconOnly, iconThreshold1..4 Type/Value
    /// - top10: rank, top10Percent, topBottom
    /// - aboveAverage: aboveBelow
    /// - timePeriod: datePeriod
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Sheet name (empty for active sheet)</param>
    /// <param name="rangeAddress">Range address (A1 notation or named range)</param>
    /// <param name="ruleType">Rule type: cellValue (or cell-value), expression, colorScale, dataBar, top10, iconSet, uniqueValues, blanksCondition, timePeriod, aboveAverage. Both camelCase and kebab-case accepted.</param>
    /// <param name="operatorType">Required for cellValue rules. XlFormatConditionOperator: equal, notEqual, greater, less, greaterEqual, lessEqual, between, notBetween</param>
    /// <param name="formula1">Required for cellValue and expression rules. First formula/value for condition</param>
    /// <param name="formula2">Required for between/notBetween cellValue rules. Second formula/value</param>
    /// <param name="interiorColor">Fill color (#RRGGBB or color index)</param>
    /// <param name="interiorPattern">Interior pattern (1=Solid, -4142=None, 9=Gray50, etc.)</param>
    /// <param name="fontColor">Font color (#RRGGBB or color index)</param>
    /// <param name="fontBold">Bold font</param>
    /// <param name="fontItalic">Italic font</param>
    /// <param name="borderStyle">Border style: none, continuous, dash, dot, etc.</param>
    /// <param name="borderColor">Border color (#RRGGBB or color index)</param>
    /// <param name="colorScaleMinType">colorScale minimum stop type: minimum, number, percent, percentile, formula</param>
    /// <param name="colorScaleMinValue">colorScale minimum stop value (for number/percent/percentile/formula)</param>
    /// <param name="colorScaleMinColor">colorScale minimum stop color (#RRGGBB)</param>
    /// <param name="colorScaleMidType">colorScale midpoint stop type (supply to create a 3-color scale)</param>
    /// <param name="colorScaleMidValue">colorScale midpoint stop value</param>
    /// <param name="colorScaleMidColor">colorScale midpoint stop color (#RRGGBB)</param>
    /// <param name="colorScaleMaxType">colorScale maximum stop type: maximum, number, percent, percentile, formula</param>
    /// <param name="colorScaleMaxValue">colorScale maximum stop value</param>
    /// <param name="colorScaleMaxColor">colorScale maximum stop color (#RRGGBB)</param>
    /// <param name="dataBarColor">dataBar fill color (#RRGGBB)</param>
    /// <param name="dataBarNegativeColor">dataBar negative-value bar color (#RRGGBB)</param>
    /// <param name="dataBarDirection">dataBar fill direction: context, leftToRight, rightToLeft</param>
    /// <param name="dataBarShowValue">dataBar show the cell value alongside the bar</param>
    /// <param name="dataBarMinType">dataBar minimum point type: automaticMinimum, minimum, number, percent, percentile, formula</param>
    /// <param name="dataBarMinValue">dataBar minimum point value</param>
    /// <param name="dataBarMaxType">dataBar maximum point type: automaticMaximum, maximum, number, percent, percentile, formula</param>
    /// <param name="dataBarMaxValue">dataBar maximum point value</param>
    /// <param name="iconSetId">iconSet id: 3Arrows, 3TrafficLights1, 4Ratings, 5Quarters, etc.</param>
    /// <param name="iconSetReverse">iconSet reverse icon order</param>
    /// <param name="iconSetShowIconOnly">iconSet show only the icon (hide the value)</param>
    /// <param name="iconThreshold1Type">iconSet threshold 1 type: percent, number, percentile, formula</param>
    /// <param name="iconThreshold1Value">iconSet threshold 1 value</param>
    /// <param name="iconThreshold2Type">iconSet threshold 2 type</param>
    /// <param name="iconThreshold2Value">iconSet threshold 2 value</param>
    /// <param name="iconThreshold3Type">iconSet threshold 3 type</param>
    /// <param name="iconThreshold3Value">iconSet threshold 3 value</param>
    /// <param name="iconThreshold4Type">iconSet threshold 4 type</param>
    /// <param name="iconThreshold4Value">iconSet threshold 4 value</param>
    /// <param name="rank">top10 rank (number of values, or percent when top10Percent is true)</param>
    /// <param name="top10Percent">top10 treat rank as a percentage</param>
    /// <param name="topBottom">top10 direction: top or bottom</param>
    /// <param name="aboveBelow">aboveAverage selector: aboveAverage, belowAverage, aboveStdDev, belowStdDev, equalAboveAverage, equalBelowAverage</param>
    /// <param name="datePeriod">timePeriod period: today, yesterday, tomorrow, last7Days, thisWeek, lastWeek, nextWeek, thisMonth, lastMonth, nextMonth</param>
    /// <exception cref="InvalidOperationException">Sheet or range not found</exception>
    /// <exception cref="ArgumentException">Invalid rule type, operator, color, or format value</exception>
    [ServiceAction("add-rule")]
    OperationResult AddRule(
        IExcelBatch batch,
        [RequiredParameter, AllowEmptyString, FromString("sheetName")] string sheetName,
        [RequiredParameter, FromString("rangeAddress")] string rangeAddress,
        [RequiredParameter, FromString("ruleType")] string ruleType,
        [FromString("operatorType")] string? operatorType,
        [FromString("formula1")] string? formula1,
        [FromString("formula2")] string? formula2,
        [FromString("interiorColor")] string? interiorColor = null,
        [FromString("interiorPattern")] string? interiorPattern = null,
        [FromString("fontColor")] string? fontColor = null,
        bool? fontBold = null,
        bool? fontItalic = null,
        [FromString("borderStyle")] string? borderStyle = null,
        [FromString("borderColor")] string? borderColor = null,
        [FromString("colorScaleMinType")] string? colorScaleMinType = null,
        [FromString("colorScaleMinValue")] string? colorScaleMinValue = null,
        [FromString("colorScaleMinColor")] string? colorScaleMinColor = null,
        [FromString("colorScaleMidType")] string? colorScaleMidType = null,
        [FromString("colorScaleMidValue")] string? colorScaleMidValue = null,
        [FromString("colorScaleMidColor")] string? colorScaleMidColor = null,
        [FromString("colorScaleMaxType")] string? colorScaleMaxType = null,
        [FromString("colorScaleMaxValue")] string? colorScaleMaxValue = null,
        [FromString("colorScaleMaxColor")] string? colorScaleMaxColor = null,
        [FromString("dataBarColor")] string? dataBarColor = null,
        [FromString("dataBarNegativeColor")] string? dataBarNegativeColor = null,
        [FromString("dataBarDirection")] string? dataBarDirection = null,
        bool? dataBarShowValue = null,
        [FromString("dataBarMinType")] string? dataBarMinType = null,
        [FromString("dataBarMinValue")] string? dataBarMinValue = null,
        [FromString("dataBarMaxType")] string? dataBarMaxType = null,
        [FromString("dataBarMaxValue")] string? dataBarMaxValue = null,
        [FromString("iconSetId")] string? iconSetId = null,
        bool? iconSetReverse = null,
        bool? iconSetShowIconOnly = null,
        [FromString("iconThreshold1Type")] string? iconThreshold1Type = null,
        [FromString("iconThreshold1Value")] string? iconThreshold1Value = null,
        [FromString("iconThreshold2Type")] string? iconThreshold2Type = null,
        [FromString("iconThreshold2Value")] string? iconThreshold2Value = null,
        [FromString("iconThreshold3Type")] string? iconThreshold3Type = null,
        [FromString("iconThreshold3Value")] string? iconThreshold3Value = null,
        [FromString("iconThreshold4Type")] string? iconThreshold4Type = null,
        [FromString("iconThreshold4Value")] string? iconThreshold4Value = null,
        int? rank = null,
        bool? top10Percent = null,
        [FromString("topBottom")] string? topBottom = null,
        [FromString("aboveBelow")] string? aboveBelow = null,
        [FromString("datePeriod")] string? datePeriod = null);

    /// <summary>
    /// Removes all conditional formatting from range
    /// Excel COM: Range.FormatConditions.Delete()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet name (empty for active sheet)</param>
    /// <param name="rangeAddress">Range address to clear rules from (e.g., A1:D10)</param>
    /// <exception cref="InvalidOperationException">Sheet or range not found</exception>
    [ServiceAction("clear-rules")]
    OperationResult ClearRules(
        IExcelBatch batch,
        [RequiredParameter, AllowEmptyString, FromString("sheetName")] string sheetName,
        [RequiredParameter, FromString("rangeAddress")] string rangeAddress);

    /// <summary>
    /// Lists existing conditional formatting rules for a range.
    /// Excel COM: Range.FormatConditions enumeration.
    /// Returns rule type, operator, formulas, applies-to range, priority, and formatting
    /// (interior/font/borders). Colors are returned as #RRGGBB hex strings. Rules are returned
    /// in priority order. Visual rule types also return their type-specific configuration:
    /// colorScale -> colorScaleCriteria, dataBar -> dataBar, iconSet -> iconSet,
    /// top10 -> top10, aboveAverage -> aboveBelow, timePeriod -> datePeriod. Each field is only
    /// present on its matching rule type.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Sheet name (empty for active sheet)</param>
    /// <param name="rangeAddress">Range address to read rules from (e.g., A1:G41)</param>
    /// <exception cref="InvalidOperationException">Sheet or range not found</exception>
    [ServiceAction("list-rules")]
    ConditionalFormatListResult ListRules(
        IExcelBatch batch,
        [RequiredParameter, AllowEmptyString, FromString("sheetName")] string sheetName,
        [RequiredParameter, FromString("rangeAddress")] string rangeAddress);

    /// <summary>
    /// Lists all conditional formatting rules on an entire worksheet.
    /// Excel COM: Worksheet.Cells.FormatConditions enumeration.
    /// Each rule includes its applies-to range so rules can be grouped by range. Colors are
    /// returned as #RRGGBB hex strings and rules are returned in priority order.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Sheet name (empty for active sheet)</param>
    /// <exception cref="InvalidOperationException">Sheet not found</exception>
    [ServiceAction("list-worksheet-rules")]
    ConditionalFormatListResult ListWorksheetRules(
        IExcelBatch batch,
        [RequiredParameter, AllowEmptyString, FromString("sheetName")] string sheetName);
}
