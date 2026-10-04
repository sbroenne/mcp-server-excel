using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConditionalFormattingTests
{
    [Fact]
    public void ListRules_NoRules_ReturnsEmptyList()
    {
        var batch = _fixture.BatchToken;

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:D10");

        Assert.True(result.Success);
        Assert.NotNull(result.Rules);
        Assert.Empty(result.Rules);
    }

    [Fact]
    public void ListRules_SingleCellValueRule_ReturnsRuleWithDetails()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "cellValue", "greater", "100", null,
            interiorColor: "#FFFF00"));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("cellValue", rule.Type);
        Assert.Equal("greater", rule.Operator);
        // Excel normalizes numeric Formula1 to a leading-'=' form ("=100").
        Assert.Equal("=100", rule.Formula1);
        Assert.Equal("#FFFF00", rule.InteriorColor);
        Assert.Equal("$A$1:$A$41", rule.AppliesTo);
        AssertNativeRule("A1:A41", Excel.XlFormatConditionType.xlCellValue, "=100");
        AssertDisplayedColor("A13", 65535);
        AssertDisplayedColor("A12", 16777215);
    }

    [Fact]
    public void ListRules_ExpressionRuleWithFontFormatting_ReturnsFontDetails()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:G41", "expression", null, "=$G1>1000", null,
            interiorColor: "#FF0000", fontColor: "#FFFFFF", fontBold: true));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:G41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("expression", rule.Type);
        Assert.Equal("=$G1>1000", rule.Formula1);
        Assert.Equal("#FF0000", rule.InteriorColor);
        Assert.Equal("#FFFFFF", rule.FontColor);
        Assert.True(rule.FontBold);
        Assert.Equal("$A$1:$G$41", rule.AppliesTo);
        AssertNativeRule("A1:G41", Excel.XlFormatConditionType.xlExpression, "=$G1>1000");
        AssertDisplayedColor("A3", 255);
        AssertDisplayedColor("A2", 16777215);
    }

    [Fact]
    public void ListRules_MultipleRules_ReturnsAllInPriorityOrder()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "cellValue", "greater", "100", null,
            interiorColor: "#FFFF00"));
        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "cellValue", "less", "0", null,
            interiorColor: "#00FF00"));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        Assert.Equal(2, result.Rules.Count);
        // cellValue rules always carry a priority, so every rule must expose one.
        Assert.All(result.Rules, r => Assert.True(r.Priority.HasValue));
        // Priorities should be in ascending collection order.
        var priorities = result.Rules
            .Select(r => r.Priority!.Value)
            .ToList();
        var sorted = priorities.OrderBy(p => p).ToList();
        Assert.Equal(sorted, priorities);
        Assert.Equal([1, 2], priorities);
        Assert.Equal(["greater", "less"], result.Rules.Select(item => item.Operator));
        Assert.Equal(["=100", "=0"], result.Rules.Select(item => item.Formula1));
        Assert.All(result.Rules, item => Assert.Equal("$A$1:$A$41", item.AppliesTo));
        AssertDisplayedColor("A1", 65280);
        AssertDisplayedColor("A13", 65535);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("#FFFFFF")]
    public void ListRules_Top10FontFlags_ReturnNativeValuesWithoutRequiringFontColor(string? fontColor)
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A10", "top10", null, null, null,
            fontColor: fontColor, fontBold: true, fontItalic: false, rank: 7, top10Percent: true));

        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.FormatConditions? conditions = null;
            Excel.Top10? nativeRule = null;
            Excel.Font? font = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                range = sheet.Range["A1:A10"];
                conditions = range.FormatConditions;
                nativeRule = (Excel.Top10)conditions.Item(1);
                font = nativeRule.Font;
                Assert.True(Assert.IsType<bool>(font.Bold));
                Assert.False(Assert.IsType<bool>(font.Italic));
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref nativeRule);
                ComUtilities.Release(ref conditions);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

        var rule = Assert.Single(RequireSuccess(
            _conditionalFormattingCommands.ListRules(batch, "", "A1:A10")).Rules);
        Assert.Equal("top10", rule.Type);
        Assert.True(rule.FontBold);
        Assert.False(rule.FontItalic);
        Assert.Equal(fontColor, rule.FontColor);
    }

    [Fact]
    public void ListWorksheetRules_AggregatesRulesAcrossRanges()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A10", "cellValue", "greater", "5", null,
            interiorColor: "#FFFF00"));
        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "C1:C10", "cellValue", "less", "5", null,
            interiorColor: "#00FF00"));

        var result = _conditionalFormattingCommands.ListWorksheetRules(batch, "");

        Assert.True(result.Success);
        Assert.Null(result.RangeAddress);
        Assert.Equal(2, result.Rules.Count);
        Assert.Equal("$A$1:$A$10", Assert.Single(result.Rules, item => item.Operator == "greater").AppliesTo);
        Assert.Equal("$C$1:$C$10", Assert.Single(result.Rules, item => item.Operator == "less").AppliesTo);
        AssertDisplayedColor("A3", 65535);
        AssertDisplayedColor("C3", 65280);
        AssertDisplayedColor("B3", 16777215);
    }

    [Fact]
    public void ListWorksheetRules_NoRules_ReturnsEmptyList()
    {
        var batch = _fixture.BatchToken;

        var result = _conditionalFormattingCommands.ListWorksheetRules(batch, "");

        Assert.True(result.Success);
        Assert.Empty(result.Rules);
    }

    [Fact]
    public void AddRule_WithBorderStyleAndColor_Succeeds()
    {
        // Regression test for #737: FormatCondition.Borders is a 4-item
        // collection indexed 1-4, not the xlEdgeLeft(7)/Top(8)/Bottom(9)/Right(10)
        // constants used for Range.Borders. Writing via those constants throws
        // COMException: "Unable to set the LineStyle property of the Border class".
        var batch = _fixture.BatchToken;

        var result = _conditionalFormattingCommands.AddRule(batch, "", "A1:A10", "cellValue", "greater", "100", null,
            borderStyle: "continuous", borderColor: "#FF0000");

        Assert.True(result.Success);
        var listed = RequireSuccess(_conditionalFormattingCommands.ListRules(batch, "", "A1:A10"));
        var rule = Assert.Single(listed.Rules);
        Assert.Equal("continuous", rule.BorderStyle);
        Assert.Equal("#FF0000", rule.BorderColor);
        AssertRuleBorders();
    }

    [Fact]
    public void AddRule_CellValueWithoutOperator_ThrowsHelpfulError()
    {
        var batch = _fixture.BatchToken;
        var before = SeedExistingRule();

        var exception = Assert.Throws<ArgumentException>(() =>
            _conditionalFormattingCommands.AddRule(batch, "", "A1:A10", "cellValue", null, "100", null));

        Assert.Contains("operatorType is required", exception.Message);
        AssertExistingRulePreserved(before);
    }

    [Fact]
    public void AddRule_CellValueWithoutFormula1_ThrowsHelpfulError()
    {
        var batch = _fixture.BatchToken;
        var before = SeedExistingRule();

        var exception = Assert.Throws<ArgumentException>(() =>
            _conditionalFormattingCommands.AddRule(batch, "", "A1:A10", "cellValue", "greater", null, null));

        Assert.Contains("formula1 is required", exception.Message);
        AssertExistingRulePreserved(before);
    }

    [Fact]
    public void AddRule_BetweenWithoutFormula2_ThrowsHelpfulError()
    {
        var batch = _fixture.BatchToken;
        var before = SeedExistingRule();

        var exception = Assert.Throws<ArgumentException>(() =>
            _conditionalFormattingCommands.AddRule(batch, "", "A1:A10", "cellValue", "between", "10", null));

        Assert.Contains("formula2 is required", exception.Message);
        AssertExistingRulePreserved(before);
    }

    [Fact]
    public void AddRule_ExpressionWithoutFormula1_ThrowsHelpfulError()
    {
        var batch = _fixture.BatchToken;
        var before = SeedExistingRule();

        var exception = Assert.Throws<ArgumentException>(() =>
            _conditionalFormattingCommands.AddRule(batch, "", "A1:A10", "expression", null, null, null));

        Assert.Contains("formula1 is required", exception.Message);
        AssertExistingRulePreserved(before);
    }

    private void AssertRuleBorders() =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.FormatConditions? conditions = null;
            Excel.FormatCondition? condition = null;
            Excel.Borders? borders = null;
            Excel.Border? border = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.ActiveSheet;
                range = sheet.Range["A1:A10"];
                conditions = range.FormatConditions;
                Assert.Equal(1, conditions.Count);
                condition = (Excel.FormatCondition)conditions.Item(1);
                borders = condition.Borders;
                Assert.Equal(4, borders.Count);
                for (int index = 1; index <= 4; index++)
                {
                    border = borders[(Excel.XlBordersIndex)index];
                    Assert.Equal((int)Excel.XlLineStyle.xlContinuous,
                        Convert.ToInt32(border.LineStyle, CultureInfo.InvariantCulture));
                    Assert.Equal(255, Convert.ToInt32(border.Color, CultureInfo.InvariantCulture));
                    ComUtilities.Release(ref border);
                }
            }
            finally
            {
                ComUtilities.Release(ref border);
                ComUtilities.Release(ref borders);
                ComUtilities.Release(ref condition);
                ComUtilities.Release(ref conditions);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

    [Fact]
    public void ListRules_RuleWithBorderStyleAndColor_RoundTripsCorrectly()
    {
        // Regression test for #737 acceptance criterion (b): border style/color
        // written via `add` must be correctly reported by `list-rules`.
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A10", "cellValue", "greater", "100", null,
            borderStyle: "continuous", borderColor: "#FF0000"));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:A10");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("continuous", rule.BorderStyle);
        Assert.Equal("#FF0000", rule.BorderColor);
        AssertRuleBorders();
    }

    [Fact]
    public void ListRules_InvalidSheet_Throws()
    {
        var batch = _fixture.BatchToken;
        var before = SeedExistingRule();

        var error = Assert.Throws<InvalidOperationException>(() =>
            _conditionalFormattingCommands.ListRules(batch, "NonExistentSheet", "A1:D10"));
        Assert.Contains("conditionalformat.list-rules failed [ComInterop/COMException]", error.Message);
        AssertExistingRulePreserved(before);
    }

    [Fact]
    public void ListRules_InvalidRange_Throws()
    {
        var batch = _fixture.BatchToken;
        var before = SeedExistingRule();

        var error = Assert.Throws<InvalidOperationException>(() =>
            _conditionalFormattingCommands.ListRules(batch, "", "NotARange!!"));
        Assert.Contains("conditionalformat.list-rules failed", error.Message);
        AssertExistingRulePreserved(before);
    }

    // === Issue #743: visual rule types expose type-specific configuration ===

    [Fact]
    public void AddRule_ColorScale_RoundTrips()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "colorScale", null, null, null,
            colorScaleMinType: "minimum", colorScaleMinColor: "#F8696B",
            colorScaleMidType: "percentile", colorScaleMidValue: "50", colorScaleMidColor: "#FFEB84",
            colorScaleMaxType: "maximum", colorScaleMaxColor: "#63BE7B"));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("colorScale", rule.Type);
        Assert.NotNull(rule.ColorScaleCriteria);
        Assert.Equal(3, rule.ColorScaleCriteria!.Count);
        Assert.Equal("minimum", rule.ColorScaleCriteria[0].Type);
        Assert.Equal("#F8696B", rule.ColorScaleCriteria[0].Color);
        Assert.Equal("percentile", rule.ColorScaleCriteria[1].Type);
        Assert.Equal("50", rule.ColorScaleCriteria[1].Value);
        Assert.Equal("#FFEB84", rule.ColorScaleCriteria[1].Color);
        Assert.Equal("maximum", rule.ColorScaleCriteria[2].Type);
        Assert.Equal("#63BE7B", rule.ColorScaleCriteria[2].Color);
        // Visual rules must not carry cellValue-only fields.
        Assert.Null(rule.DataBar);
        Assert.Null(rule.IconSet);
    }

    [Fact]
    public void AddRule_DataBar_RoundTrips()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "dataBar", null, null, null,
            dataBarColor: "#638EC6", dataBarDirection: "leftToRight", dataBarShowValue: true,
            dataBarMinType: "minimum", dataBarMaxType: "maximum"));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("dataBar", rule.Type);
        Assert.NotNull(rule.DataBar);
        Assert.Equal("#638EC6", rule.DataBar!.FillColor);
        Assert.Equal("leftToRight", rule.DataBar.Direction);
        Assert.True(rule.DataBar.ShowValue);
        Assert.Equal("minimum", rule.DataBar.MinType);
        Assert.Equal("maximum", rule.DataBar.MaxType);
        Assert.Null(rule.ColorScaleCriteria);
    }

    [Fact]
    public void AddRule_IconSet_RoundTrips()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "iconSet", null, null, null,
            iconSetId: "3TrafficLights1", iconSetReverse: false, iconSetShowIconOnly: false,
            iconThreshold1Type: "percent", iconThreshold1Value: "33",
            iconThreshold2Type: "percent", iconThreshold2Value: "67"));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("iconSet", rule.Type);
        Assert.NotNull(rule.IconSet);
        Assert.Equal("3TrafficLights1", rule.IconSet!.Id);
        Assert.False(rule.IconSet.Reverse);
        Assert.False(rule.IconSet.ShowIconOnly);
        Assert.NotNull(rule.IconSet.Criteria);
        Assert.Equal(3, rule.IconSet.Criteria!.Count);
        Assert.Equal("percent", rule.IconSet.Criteria[1].Type);
        Assert.Equal("33", rule.IconSet.Criteria[1].Value);
        Assert.Equal("percent", rule.IconSet.Criteria[2].Type);
        Assert.Equal("67", rule.IconSet.Criteria[2].Value);
        Assert.Null(rule.DataBar);
    }

    [Fact]
    public void AddRule_Top10_RoundTrips()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "top10", null, null, null,
            rank: 10, top10Percent: false, topBottom: "top",
            interiorColor: "#FFC7CE"));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("top10", rule.Type);
        Assert.NotNull(rule.Top10);
        Assert.Equal(10, rule.Top10!.Rank);
        Assert.False(rule.Top10.Percent);
        Assert.Equal("top", rule.Top10.TopBottom);
        Assert.Equal("#FFC7CE", rule.InteriorColor);
    }

    [Fact]
    public void AddRule_AboveAverage_RoundTrips()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "aboveAverage", null, null, null,
            aboveBelow: "belowAverage", interiorColor: "#FFEB9C"));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("aboveAverage", rule.Type);
        Assert.Equal("belowAverage", rule.AboveBelow);
        Assert.Equal("#FFEB9C", rule.InteriorColor);
    }

    [Fact]
    public void AddRule_TimePeriod_RoundTrips()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "timePeriod", null, null, null,
            datePeriod: "last7Days", interiorColor: "#C6EFCE"));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("timePeriod", rule.Type);
        Assert.Equal("last7Days", rule.DatePeriod);
        Assert.Equal("#C6EFCE", rule.InteriorColor);
    }

    [Fact]
    public void ListRules_MixedRuleTypes_EachHasOnlyItsFields()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "cellValue", "greater", "100", null,
            interiorColor: "#FFFF00"));
        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "colorScale", null, null, null,
            colorScaleMinType: "minimum", colorScaleMinColor: "#F8696B",
            colorScaleMaxType: "maximum", colorScaleMaxColor: "#63BE7B"));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        Assert.Equal(2, result.Rules.Count);

        var cellValue = Assert.Single(result.Rules, r => r.Type == "cellValue");
        Assert.Null(cellValue.ColorScaleCriteria);
        Assert.Null(cellValue.DataBar);
        Assert.Null(cellValue.IconSet);
        Assert.Null(cellValue.Top10);

        var colorScale = Assert.Single(result.Rules, r => r.Type == "colorScale");
        Assert.NotNull(colorScale.ColorScaleCriteria);
        Assert.Null(colorScale.Operator);
        Assert.Null(colorScale.DataBar);
    }

    [Fact]
    public void ListRules_CellValueRule_HasNoVisualFields()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "A1:A41", "cellValue", "between", "10", "20",
            interiorColor: "#FFFF00"));

        var result = _conditionalFormattingCommands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("cellValue", rule.Type);
        Assert.Null(rule.ColorScaleCriteria);
        Assert.Null(rule.DataBar);
        Assert.Null(rule.IconSet);
        Assert.Null(rule.Top10);
        Assert.Null(rule.AboveBelow);
        Assert.Null(rule.DatePeriod);
        Assert.Equal("between", rule.Operator);
        Assert.Equal("=10", rule.Formula1);
        Assert.Equal("=20", rule.Formula2);
        AssertDisplayedColor("A3", 65535);
        AssertDisplayedColor("A2", 16777215);
    }

    private string SeedExistingRule()
    {
        RequireSuccess(_conditionalFormattingCommands.AddRule(_fixture.BatchToken, "", "A1:A10",
            "cellValue", "greater", "5", null, interiorColor: "#FFFF00"));
        AssertDisplayedColor("A3", 65535);
        return JsonSerializer.Serialize(RequireSuccess(
            _conditionalFormattingCommands.ListRules(_fixture.BatchToken, "", "A1:A10")).Rules);
    }

    [Fact]
    public void ClearRules_RemovesOnlyTargetRulesAndPreservesCells()
    {
        var batch = _fixture.BatchToken;
        SeedExistingRule();
        RequireSuccess(_conditionalFormattingCommands.AddRule(batch, "", "C1:C10",
            "cellValue", "less", "5", null, interiorColor: "#00FF00"));
        AssertDisplayedColor("C3", 65280);
        var before = RequireSuccess(_rangeCommands.GetValues(batch, _sheetName, "A1:G41"));

        RequireSuccess(_conditionalFormattingCommands.ClearRules(batch, "", "A1:A10"));

        Assert.Empty(RequireSuccess(_conditionalFormattingCommands.ListRules(batch, "", "A1:A10")).Rules);
        var remaining = Assert.Single(RequireSuccess(
            _conditionalFormattingCommands.ListWorksheetRules(batch, "")).Rules);
        Assert.Equal("$C$1:$C$10", remaining.AppliesTo);
        AssertNativeRule("C1:C10", Excel.XlFormatConditionType.xlCellValue, "=5");
        AssertDisplayedColor("A3", 16777215);
        AssertDisplayedColor("C3", 65280);
        Assert.Equal(JsonSerializer.Serialize(before.Values), JsonSerializer.Serialize(
            RequireSuccess(_rangeCommands.GetValues(batch, _sheetName, "A1:G41")).Values));
    }

    [Fact]
    public void ClearRules_InvalidRange_PreservesExistingRules()
    {
        var before = SeedExistingRule();

        var error = Assert.Throws<InvalidOperationException>(() =>
            _conditionalFormattingCommands.ClearRules(_fixture.BatchToken, "", "NotARange!!"));

        Assert.Contains("conditionalformat.clear-rules failed", error.Message);
        AssertExistingRulePreserved(before);
    }

    [Theory]
    [InlineData("interior")]
    [InlineData("font")]
    [InlineData("border")]
    [InlineData("scale-min")]
    [InlineData("scale-mid")]
    [InlineData("scale-max")]
    [InlineData("bar-fill")]
    [InlineData("bar-negative")]
    public void AddRule_InvalidColor_PreservesExistingRules(string target)
    {
        var before = SeedExistingRule();

        var ruleType = target.StartsWith("scale-", StringComparison.Ordinal) ? "colorScale"
            : target.StartsWith("bar-", StringComparison.Ordinal) ? "dataBar" : "cellValue";
        var error = Assert.Throws<ArgumentException>(() =>
            _conditionalFormattingCommands.AddRule(_fixture.BatchToken, "", "A1:A10",
                ruleType, "less", "0", null,
                interiorColor: target == "interior" ? "not-a-color" : "#FF0000",
                fontColor: target == "font" ? "not-a-color" : "#FFFFFF",
                borderColor: target == "border" ? "not-a-color" : "#00FF00",
                colorScaleMinColor: target == "scale-min" ? "not-a-color" : "#FF0000",
                colorScaleMidType: "percentile", colorScaleMidValue: "50",
                colorScaleMidColor: target == "scale-mid" ? "not-a-color" : "#FFFF00",
                colorScaleMaxColor: target == "scale-max" ? "not-a-color" : "#00FF00",
                dataBarColor: target == "bar-fill" ? "not-a-color" : "#FF0000",
                dataBarNegativeColor: target == "bar-negative" ? "not-a-color" : "#00FF00"));

        Assert.Contains("color", error.Message, StringComparison.OrdinalIgnoreCase);
        AssertExistingRulePreserved(before);
    }

    private void AssertExistingRulePreserved(string before)
    {
        Assert.Equal(before, JsonSerializer.Serialize(RequireSuccess(
            _conditionalFormattingCommands.ListRules(_fixture.BatchToken, "", "A1:A10")).Rules));
        AssertNativeRule("A1:A10", Excel.XlFormatConditionType.xlCellValue, "=5");
        AssertDisplayedColor("A3", 65535);
        var values = RequireSuccess(_rangeCommands.GetValues(_fixture.BatchToken, _sheetName, "A1:G41"));
        Assert.Equal(41, values.Values.Count);
        for (var index = 0; index < 41; index++)
        {
            Assert.Equal(7, values.Values[index].Count);
            Assert.Equal(index * 10 - 10, Convert.ToInt32(values.Values[index][0], CultureInfo.InvariantCulture));
            Assert.Equal(index + 999, Convert.ToInt32(values.Values[index][6], CultureInfo.InvariantCulture));
            Assert.Equal([1, 2, 3, 4, 5], values.Values[index].Skip(1).Take(5)
                .Select(value => Convert.ToInt32(value, CultureInfo.InvariantCulture)));
        }
    }

    private void AssertNativeRule(string address, Excel.XlFormatConditionType type, string formula)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.FormatConditions? conditions = null;
            Excel.FormatCondition? condition = null;
            Excel.Range? appliesTo = null;
            try
            {
                sheet = (Excel.Worksheet)context.Book.ActiveSheet;
                range = sheet.Range[address];
                conditions = range.FormatConditions;
                Assert.Equal(1, conditions.Count);
                condition = (Excel.FormatCondition)conditions.Item(1);
                Assert.Equal((int)type, condition.Type);
                Assert.Equal(formula, condition.Formula1);
                appliesTo = condition.AppliesTo;
                Assert.Equal(range.Address, appliesTo.Address);
            }
            finally
            {
                ComUtilities.Release(ref appliesTo);
                ComUtilities.Release(ref condition);
                ComUtilities.Release(ref conditions);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private void AssertDisplayedColor(string address, int color)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            Excel.DisplayFormat? format = null;
            Excel.Interior? interior = null;
            try
            {
                sheet = (Excel.Worksheet)context.Book.ActiveSheet;
                cell = sheet.Range[address];
                format = cell.DisplayFormat;
                interior = format.Interior;
                Assert.Equal(color, Convert.ToInt32(interior.Color, CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref format);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
