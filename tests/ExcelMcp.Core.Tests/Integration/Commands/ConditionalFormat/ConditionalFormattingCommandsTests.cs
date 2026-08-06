using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.ConditionalFormat;

/// <summary>
/// Integration tests for ConditionalFormattingCommands read operations
/// (list-rules / list-worksheet-rules). Exercises real Excel COM automation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "ConditionalFormat")]
[Trait("RequiresExcel", "true")]
public class ConditionalFormattingCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly ConditionalFormattingCommands _commands;
    private readonly TempDirectoryFixture _fixture;

    /// <summary>
    /// Initializes a new instance of the test class.
    /// </summary>
    public ConditionalFormattingCommandsTests(TempDirectoryFixture fixture)
    {
        _commands = new ConditionalFormattingCommands();
        _fixture = fixture;
    }

    [Fact]
    public void ListRules_NoRules_ReturnsEmptyList()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        var result = _commands.ListRules(batch, "", "A1:D10");

        Assert.True(result.Success);
        Assert.NotNull(result.Rules);
        Assert.Empty(result.Rules);
    }

    [Fact]
    public void ListRules_SingleCellValueRule_ReturnsRuleWithDetails()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A41", "cellValue", "greater", "100", null,
            interiorColor: "#FFFF00");

        var result = _commands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("cellValue", rule.Type);
        Assert.Equal("greater", rule.Operator);
        // Excel normalizes numeric Formula1 to a leading-'=' form ("=100").
        Assert.Equal("=100", rule.Formula1);
        Assert.Equal("#FFFF00", rule.InteriorColor);
        Assert.False(string.IsNullOrEmpty(rule.AppliesTo));
    }

    [Fact]
    public void ListRules_ExpressionRuleWithFontFormatting_ReturnsFontDetails()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:G41", "expression", null, "=$G1>1000", null,
            interiorColor: "#FF0000", fontColor: "#FFFFFF", fontBold: true);

        var result = _commands.ListRules(batch, "", "A1:G41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("expression", rule.Type);
        Assert.Equal("=$G1>1000", rule.Formula1);
        Assert.Equal("#FF0000", rule.InteriorColor);
        Assert.Equal("#FFFFFF", rule.FontColor);
        Assert.True(rule.FontBold);
    }

    [Fact]
    public void ListRules_MultipleRules_ReturnsAllInPriorityOrder()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A41", "cellValue", "greater", "100", null,
            interiorColor: "#FFFF00");
        _commands.AddRule(batch, "", "A1:A41", "cellValue", "less", "0", null,
            interiorColor: "#00FF00");

        var result = _commands.ListRules(batch, "", "A1:A41");

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
    }

    [Fact]
    public void ListWorksheetRules_AggregatesRulesAcrossRanges()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A10", "cellValue", "greater", "5", null,
            interiorColor: "#FFFF00");
        _commands.AddRule(batch, "", "C1:C10", "cellValue", "less", "5", null,
            interiorColor: "#00FF00");

        var result = _commands.ListWorksheetRules(batch, "");

        Assert.True(result.Success);
        Assert.Null(result.RangeAddress);
        Assert.True(result.Rules.Count >= 2);
        Assert.All(result.Rules, r => Assert.False(string.IsNullOrEmpty(r.AppliesTo)));
    }

    [Fact]
    public void ListWorksheetRules_NoRules_ReturnsEmptyList()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        var result = _commands.ListWorksheetRules(batch, "");

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
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        var result = _commands.AddRule(batch, "", "A1:A10", "cellValue", "greater", "100", null,
            borderStyle: "continuous", borderColor: "#FF0000");

        Assert.True(result.Success);
    }

    [Fact]
    public void AddRule_CellValueWithoutOperator_ThrowsHelpfulError()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.AddRule(batch, "", "A1:A10", "cellValue", null, "100", null));

        Assert.Contains("operatorType is required", exception.Message);
    }

    [Fact]
    public void AddRule_CellValueWithoutFormula1_ThrowsHelpfulError()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.AddRule(batch, "", "A1:A10", "cellValue", "greater", null, null));

        Assert.Contains("formula1 is required", exception.Message);
    }

    [Fact]
    public void AddRule_BetweenWithoutFormula2_ThrowsHelpfulError()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.AddRule(batch, "", "A1:A10", "cellValue", "between", "10", null));

        Assert.Contains("formula2 is required", exception.Message);
    }

    [Fact]
    public void AddRule_ExpressionWithoutFormula1_ThrowsHelpfulError()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.AddRule(batch, "", "A1:A10", "expression", null, null, null));

        Assert.Contains("formula1 is required", exception.Message);
    }

    [Fact]
    public void ListRules_RuleWithBorderStyleAndColor_RoundTripsCorrectly()
    {
        // Regression test for #737 acceptance criterion (b): border style/color
        // written via `add` must be correctly reported by `list-rules`.
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A10", "cellValue", "greater", "100", null,
            borderStyle: "continuous", borderColor: "#FF0000");

        var result = _commands.ListRules(batch, "", "A1:A10");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("continuous", rule.BorderStyle);
        Assert.Equal("#FF0000", rule.BorderColor);
    }

    [Fact]
    public void ListRules_InvalidSheet_Throws()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        Assert.ThrowsAny<Exception>(() =>
            _commands.ListRules(batch, "NonExistentSheet", "A1:D10"));
    }

    [Fact]
    public void ListRules_InvalidRange_Throws()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        Assert.ThrowsAny<Exception>(() =>
            _commands.ListRules(batch, "", "NotARange!!"));
    }

    // === Issue #743: visual rule types expose type-specific configuration ===

    [Fact]
    public void AddRule_ColorScale_RoundTrips()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A41", "colorScale", null, null, null,
            colorScaleMinType: "minimum", colorScaleMinColor: "#F8696B",
            colorScaleMidType: "percentile", colorScaleMidValue: "50", colorScaleMidColor: "#FFEB84",
            colorScaleMaxType: "maximum", colorScaleMaxColor: "#63BE7B");

        var result = _commands.ListRules(batch, "", "A1:A41");

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
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A41", "dataBar", null, null, null,
            dataBarColor: "#638EC6", dataBarDirection: "leftToRight", dataBarShowValue: true,
            dataBarMinType: "minimum", dataBarMaxType: "maximum");

        var result = _commands.ListRules(batch, "", "A1:A41");

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
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A41", "iconSet", null, null, null,
            iconSetId: "3TrafficLights1", iconSetReverse: false, iconSetShowIconOnly: false,
            iconThreshold1Type: "percent", iconThreshold1Value: "33",
            iconThreshold2Type: "percent", iconThreshold2Value: "67");

        var result = _commands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("iconSet", rule.Type);
        Assert.NotNull(rule.IconSet);
        Assert.Equal("3TrafficLights1", rule.IconSet!.Id);
        Assert.False(rule.IconSet.Reverse);
        Assert.False(rule.IconSet.ShowIconOnly);
        Assert.NotNull(rule.IconSet.Criteria);
        Assert.True(rule.IconSet.Criteria!.Count >= 2);
        Assert.Null(rule.DataBar);
    }

    [Fact]
    public void AddRule_Top10_RoundTrips()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A41", "top10", null, null, null,
            rank: 10, top10Percent: false, topBottom: "top",
            interiorColor: "#FFC7CE");

        var result = _commands.ListRules(batch, "", "A1:A41");

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
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A41", "aboveAverage", null, null, null,
            aboveBelow: "belowAverage", interiorColor: "#FFEB9C");

        var result = _commands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("aboveAverage", rule.Type);
        Assert.Equal("belowAverage", rule.AboveBelow);
        Assert.Equal("#FFEB9C", rule.InteriorColor);
    }

    [Fact]
    public void AddRule_TimePeriod_RoundTrips()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A41", "timePeriod", null, null, null,
            datePeriod: "last7Days", interiorColor: "#C6EFCE");

        var result = _commands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("timePeriod", rule.Type);
        Assert.Equal("last7Days", rule.DatePeriod);
        Assert.Equal("#C6EFCE", rule.InteriorColor);
    }

    [Fact]
    public void ListRules_MixedRuleTypes_EachHasOnlyItsFields()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A41", "cellValue", "greater", "100", null,
            interiorColor: "#FFFF00");
        _commands.AddRule(batch, "", "A1:A41", "colorScale", null, null, null,
            colorScaleMinType: "minimum", colorScaleMinColor: "#F8696B",
            colorScaleMaxType: "maximum", colorScaleMaxColor: "#63BE7B");

        var result = _commands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        Assert.Equal(2, result.Rules.Count);

        var cellValue = result.Rules.Single(r => r.Type == "cellValue");
        Assert.Null(cellValue.ColorScaleCriteria);
        Assert.Null(cellValue.DataBar);
        Assert.Null(cellValue.IconSet);
        Assert.Null(cellValue.Top10);

        var colorScale = result.Rules.Single(r => r.Type == "colorScale");
        Assert.NotNull(colorScale.ColorScaleCriteria);
        Assert.Null(colorScale.Operator);
        Assert.Null(colorScale.DataBar);
    }

    [Fact]
    public void ListRules_CellValueRule_HasNoVisualFields()
    {
        var file = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(file);

        _commands.AddRule(batch, "", "A1:A41", "cellValue", "between", "10", "20",
            interiorColor: "#FFFF00");

        var result = _commands.ListRules(batch, "", "A1:A41");

        Assert.True(result.Success);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("cellValue", rule.Type);
        Assert.Null(rule.ColorScaleCriteria);
        Assert.Null(rule.DataBar);
        Assert.Null(rule.IconSet);
        Assert.Null(rule.Top10);
        Assert.Null(rule.AboveBelow);
        Assert.Null(rule.DatePeriod);
    }
}
