using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Option checks that must happen before Excel creates a conditional formatting rule,
/// and the error reported when Excel rejects a value after the rule exists.
/// </summary>
public sealed partial class PersistentServiceConditionalFormattingTests
{
    [Theory]
    [InlineData("scale-type", "Invalid threshold type")]
    [InlineData("bar-type", "Invalid threshold type")]
    [InlineData("bar-direction", "Invalid data bar direction")]
    [InlineData("icon-set", "Invalid icon set id")]
    [InlineData("icon-type", "Invalid threshold type")]
    [InlineData("top-bottom", "Invalid topBottom value")]
    [InlineData("rank-zero", "rank")]
    [InlineData("rank-too-big", "rank")]
    [InlineData("percent-rank-too-big", "rank")]
    [InlineData("above-below", "Invalid aboveBelow value")]
    [InlineData("pattern", "Unknown interior pattern")]
    [InlineData("border-style", "border")]
    public void AddRule_InvalidOption_RejectsBeforeCreatingRule(string target, string expectedMessage)
    {
        var before = SeedExistingRule();
        var ruleType = target switch
        {
            "scale-type" => "colorScale",
            "bar-type" or "bar-direction" => "dataBar",
            "icon-set" or "icon-type" => "iconSet",
            "top-bottom" or "rank-zero" or "rank-too-big" or "percent-rank-too-big" => "top10",
            "above-below" => "aboveAverage",
            _ => "cellValue",
        };

        var error = Assert.Throws<ArgumentException>(() =>
            _conditionalFormattingCommands.AddRule(_fixture.BatchToken, "", "A1:A10",
                ruleType, "less", "0", null,
                interiorPattern: target == "pattern" ? "zigzag" : null,
                borderStyle: target == "border-style" ? "wavy" : null,
                colorScaleMinType: target == "scale-type" ? "bogus" : null,
                dataBarMinType: target == "bar-type" ? "bogus" : null,
                dataBarDirection: target == "bar-direction" ? "sideways" : null,
                iconSetId: target == "icon-set" ? "7Smileys" : null,
                iconThreshold1Type: target == "icon-type" ? "bogus" : null,
                topBottom: target == "top-bottom" ? "middle" : null,
                rank: target switch
                {
                    "rank-zero" => 0,
                    "rank-too-big" => 1001,
                    "percent-rank-too-big" => 101,
                    _ => null,
                },
                top10Percent: target == "percent-rank-too-big" ? true : null,
                aboveBelow: target == "above-below" ? "sideways" : null));

        Assert.Contains(expectedMessage, error.Message, StringComparison.OrdinalIgnoreCase);
        AssertExistingRulePreserved(before);
    }

    [Theory]
    [InlineData("colorScale")]
    [InlineData("cellValue")]
    public void AddRule_ValueRejectedAfterRuleExists_ReportsLeftoverRule(string ruleType)
    {
        var before = SeedExistingRule();

        var error = Assert.Throws<InvalidOperationException>(() =>
            _conditionalFormattingCommands.AddRule(_fixture.BatchToken, "", "A1:A10",
                ruleType, "less", "0", null,
                interiorPattern: ruleType == "cellValue" ? "999" : null,
                colorScaleMinType: ruleType == "colorScale" ? "formula" : null,
                colorScaleMinValue: ruleType == "colorScale" ? "=SUM(" : null));

        Assert.Contains($"{ruleType} rule", error.Message, StringComparison.Ordinal);
        Assert.Contains("'A1:A10'", error.Message, StringComparison.Ordinal);
        Assert.Contains($"sheet '{_sheetName}'", error.Message, StringComparison.Ordinal);
        Assert.Contains("remains", error.Message, StringComparison.Ordinal);
        var rules = RequireSuccess(_conditionalFormattingCommands.ListRules(_fixture.BatchToken, "", "A1:A10")).Rules;
        Assert.Equal(2, rules.Count);
        Assert.Contains(rules, rule => rule.Type == ruleType && JsonSerializer.Serialize(rule) != before);
    }
}
