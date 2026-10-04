using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "ConditionalRuleEditing")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceConditionalRuleEditingTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IConditionalFormattingCommands _conditional =
        fixture.CreateCommands<IConditionalFormattingCommands>();

    [Fact]
    public void UpdateSelectedRule_PreservesOtherRulesAndAppliesExactNativeScope()
    {
        var sheet = CreateRules();
        var first = ReadRule(sheet, "=10");
        var unrelated = ReadRule(sheet, "=20");
        var response = _fixture.Send("conditionalformat.update-rule", new
        {
            sheetName = sheet,
            rulePriority = first.priority,
            expectedFingerprint = first.fingerprint,
            options = new { formula1 = "15", appliesTo = "A1:A2,A4:A5", stopIfTrue = true }
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var after = ReadRule(sheet, "=15");
        Assert.NotEqual(first.fingerprint, after.fingerprint);
        Assert.Equal(unrelated.fingerprint, ReadRule(sheet, "=20").fingerprint);
        var listed = _conditional.ListWorksheetRules(_fixture.BatchToken, sheet);
        Assert.True(listed.Success, listed.ErrorMessage);
        var updated = Assert.Single(listed.Rules, rule => rule.Formula1 == "=15");
        Assert.True(updated.StopIfTrue);
        Assert.Equal("$A$1:$A$2,$A$4:$A$5", updated.AppliesTo);
    }

    [Fact]
    public async Task StaleFingerprint_RejectsDeletionAfterRuleChanged()
    {
        var sheet = CreateRules();
        var first = ReadRule(sheet, "=10");
        var unrelated = ReadRule(sheet, "=20");
        var updated = _fixture.Send("conditionalformat.update-rule", new
        {
            sheetName = sheet,
            rulePriority = first.priority,
            expectedFingerprint = first.fingerprint,
            options = new { formula1 = "15" }
        });
        using var updateResult = JsonDocument.Parse(updated.Result!);
        Assert.True(updateResult.RootElement.GetProperty("success").GetBoolean());
        var rejected = await _fixture.SendForFailureAsync("conditionalformat.delete-rule", new
        {
            sheetName = sheet,
            rulePriority = first.priority,
            expectedFingerprint = first.fingerprint
        });
        Assert.False(rejected.Success);
        Assert.Contains("changed", rejected.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var listed = _conditional.ListWorksheetRules(_fixture.BatchToken, sheet);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal(2, listed.Rules.Count);
        Assert.Contains(listed.Rules, rule => rule.Formula1 == "=15");
        Assert.Equal(unrelated, ReadRule(sheet, "=20"));
    }

    [Fact]
    public void ReorderThenDeleteSelectedRule_PreservesUnrelatedRule()
    {
        var sheet = CreateRules();
        var first = ReadRule(sheet, "=10");
        int targetPriority = first.priority == 1 ? 2 : 1;
        var reordered = _fixture.Send("conditionalformat.set-rule-priority", new
        {
            sheetName = sheet,
            rulePriority = first.priority,
            expectedFingerprint = first.fingerprint,
            newPriority = targetPriority
        });
        using var reorderResult = JsonDocument.Parse(reordered.Result!);
        Assert.True(reorderResult.RootElement.GetProperty("success").GetBoolean());
        var current = ReadRule(sheet, "=10");
        Assert.Equal(targetPriority, current.priority);
        Assert.NotEqual(current.priority, ReadRule(sheet, "=20").priority);
        var deleted = _fixture.Send("conditionalformat.delete-rule", new
        {
            sheetName = sheet,
            rulePriority = current.priority,
            expectedFingerprint = current.fingerprint
        });
        using var deleteResult = JsonDocument.Parse(deleted.Result!);
        Assert.True(deleteResult.RootElement.GetProperty("success").GetBoolean());
        var listed = _conditional.ListWorksheetRules(_fixture.BatchToken, sheet);
        Assert.True(listed.Success, listed.ErrorMessage);
        var remaining = Assert.Single(listed.Rules);
        Assert.Equal("=20", remaining.Formula1);
        Assert.Equal("$C$1:$C$5", remaining.AppliesTo);
        Assert.Equal("#00FF00", remaining.InteriorColor);
    }

    [Fact]
    public void UpdateFormulaAndStopIfTrue_PreservesNativeRule()
    {
        var sheet = CreateRules();
        var first = ReadRule(sheet, "=10");
        var response = _fixture.Send("conditionalformat.update-rule", new
        {
            sheetName = sheet,
            rulePriority = first.priority,
            expectedFingerprint = first.fingerprint,
            options = new { formula1 = "15", stopIfTrue = true }
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var listed = _conditional.ListWorksheetRules(_fixture.BatchToken, sheet);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.True(Assert.Single(listed.Rules, rule => rule.Formula1 == "=15").StopIfTrue);
    }

    [Fact]
    public void UpdateDisjointAppliesToOnly_PreservesNativeRule()
    {
        var sheet = CreateRules();
        var first = ReadRule(sheet, "=10");
        var response = _fixture.Send("conditionalformat.update-rule", new
        {
            sheetName = sheet,
            rulePriority = first.priority,
            expectedFingerprint = first.fingerprint,
            options = new { appliesTo = "A1:A2,A4:A5" }
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var listed = _conditional.ListWorksheetRules(_fixture.BatchToken, sheet);
        Assert.True(listed.Success, listed.ErrorMessage);
        var updated = Assert.Single(listed.Rules, rule => rule.Formula1 == "=10");
        Assert.Equal("$A$1:$A$2,$A$4:$A$5", updated.AppliesTo);
    }

    [Fact]
    public void AddRule_ExplicitStopIfTrueAndPriority_AreApplied()
    {
        var sheet = CreateRules();
        var response = _fixture.Send("conditionalformat.add-rule", new
        {
            sheetName = sheet,
            rangeAddress = "E1:E5",
            ruleType = "cellValue",
            operatorType = "greater",
            formula1 = "30",
            priority = 1,
            stopIfTrue = false
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var listed = _conditional.ListWorksheetRules(_fixture.BatchToken, sheet);
        Assert.True(listed.Success, listed.ErrorMessage);
        var added = Assert.Single(listed.Rules, rule => rule.Formula1 == "=30");
        Assert.Equal(1, added.Priority);
        Assert.False(added.StopIfTrue);
        Assert.Equal(3, listed.Rules.Count);
    }

    [Theory]
    [InlineData("colorScale", """{"colorScaleMinType":"percentile","colorScaleMinValue":"25","colorScaleMinColor":"#123456"}""")]
    [InlineData("dataBar", """{"dataBarColor":"#123456","dataBarNegativeColor":"#654321","dataBarMinType":"number","dataBarMinValue":"-10","dataBarShowValue":false}""")]
    [InlineData("iconSet", """{"iconSetId":"4Ratings","iconSetReverse":true,"iconSetShowIconOnly":true,"iconThreshold3Type":"percent","iconThreshold3Value":"80"}""")]
    [InlineData("top10", """{"rank":5,"top10Percent":false,"topBottom":"bottom","stopIfTrue":false}""")]
    [InlineData("aboveAverage", """{"aboveBelow":"belowStdDev","standardDeviations":2}""")]
    [InlineData("timePeriod", """{"datePeriod":"tomorrow"}""")]
    [InlineData("uniqueValues", """{"duplicateValues":true}""")]
    public void VisualRuleUpdates_RetainTypeAndReadNativeSettings(string type, string optionsJson)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var added = _conditional.AddRule(_fixture.BatchToken, sheet, "A1:A5", type,
            null, null, null, datePeriod: type == "timePeriod" ? "today" : null);
        Assert.True(added.Success, added.ErrorMessage);
        var before = _conditional.ListWorksheetRules(_fixture.BatchToken, sheet);
        Assert.True(before.Success, before.ErrorMessage);
        var original = Assert.Single(before.Rules);
        using var options = JsonDocument.Parse(optionsJson);
        var response = _fixture.Send("conditionalformat.update-rule", new
        {
            sheetName = sheet,
            rulePriority = original.Priority,
            expectedFingerprint = original.Fingerprint,
            options = options.RootElement
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var after = _conditional.ListWorksheetRules(_fixture.BatchToken, sheet);
        Assert.True(after.Success, after.ErrorMessage);
        var rule = Assert.Single(after.Rules);
        Assert.Equal(type, rule.Type);
        Assert.NotEqual(original.Fingerprint, rule.Fingerprint);
        switch (type)
        {
            case "colorScale":
                var stop = Assert.IsType<List<Sbroenne.ExcelMcp.Core.Models.ColorScaleCriterionInfo>>(rule.ColorScaleCriteria)[0];
                Assert.Equal("percentile", stop.Type);
                Assert.Equal("25", stop.Value);
                Assert.Equal("#123456", stop.Color);
                break;
            case "dataBar":
                Assert.NotNull(rule.DataBar);
                Assert.Equal("#123456", rule.DataBar.FillColor);
                Assert.Equal("#654321", rule.DataBar.BarColorNegative);
                Assert.Equal("number", rule.DataBar.MinType);
                Assert.Equal("-10", rule.DataBar.MinValue);
                Assert.False(rule.DataBar.ShowValue);
                break;
            case "iconSet":
                Assert.NotNull(rule.IconSet);
                Assert.Equal("4Ratings", rule.IconSet.Id);
                Assert.True(rule.IconSet.Reverse);
                Assert.True(rule.IconSet.ShowIconOnly);
                Assert.NotNull(rule.IconSet.Criteria);
                Assert.Equal(4, rule.IconSet.Criteria.Count);
                Assert.Equal("80", rule.IconSet.Criteria[3].Value);
                Assert.Equal("percent", rule.IconSet.Criteria[3].Type);
                break;
            case "top10":
                Assert.NotNull(rule.Top10);
                Assert.Equal(5, rule.Top10.Rank);
                Assert.Equal("bottom", rule.Top10.TopBottom);
                Assert.False(rule.Top10.Percent);
                Assert.False(rule.StopIfTrue);
                break;
            case "aboveAverage":
                Assert.Equal("belowStdDev", rule.AboveBelow);
                Assert.Equal(2, rule.StandardDeviations);
                break;
            case "timePeriod": Assert.Equal("tomorrow", rule.DatePeriod); break;
            case "uniqueValues": Assert.True(rule.DuplicateValues); break;
        }
    }

    [Theory]
    [InlineData("""{"dataBarColor":"#123456"}""")]
    [InlineData("""{"interiorColor":"invalid","formula1":"15"}""")]
    [InlineData("""{"operatorType":"between","formula2":null}""")]
    [InlineData("""{"misspelledSetting":true}""")]
    [InlineData("""{}""")]
    public async Task InvalidOptions_DoNotChangeSelectedOrOtherRule(string optionsJson)
    {
        var sheet = CreateRules();
        var first = ReadRule(sheet, "=10");
        var other = ReadRule(sheet, "=20");
        using var options = JsonDocument.Parse(optionsJson);
        var failed = await _fixture.SendForFailureAsync("conditionalformat.update-rule", new
        {
            sheetName = sheet,
            rulePriority = first.priority,
            expectedFingerprint = first.fingerprint,
            options = options.RootElement
        });
        Assert.False(failed.Success);
        Assert.False(string.IsNullOrWhiteSpace(failed.ErrorMessage));
        Assert.Equal(first, ReadRule(sheet, "=10"));
        Assert.Equal(other, ReadRule(sheet, "=20"));
    }

    [Theory]
    [InlineData("aboveAverage", "00011")]
    [InlineData("belowAverage", "11000")]
    [InlineData("equalAboveAverage", "00111")]
    [InlineData("equalBelowAverage", "11100")]
    [InlineData("aboveStdDev", "00001")]
    [InlineData("belowStdDev", "10000")]
    public void AverageComparisons_ApplyTheRequestedNativeMeaning(string comparison, string highlighted)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:A5",
            [[1], [2], [3], [4], [5]]).Success);
        var added = _conditional.AddRule(_fixture.BatchToken, sheet, "A1:A5", "aboveAverage",
            null, null, null, interiorColor: "#123456", aboveBelow: comparison);
        Assert.True(added.Success, added.ErrorMessage);
        var formatted = _fixture.Send("rangeformat.get-format", new
        {
            sheetName = sheet,
            rangeAddress = "A1:A5",
            view = "Displayed"
        });
        using var result = JsonDocument.Parse(formatted.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var cells = result.RootElement.GetProperty("cells").EnumerateArray().ToArray();
        Assert.Equal(5, cells.Length);
        for (int i = 0; i < cells.Length; i++)
        {
            var fill = cells[i].GetProperty("displayed").GetProperty("fill");
            bool colored = fill.TryGetProperty("color", out var color) &&
                color.TryGetProperty("rgb", out var rgb) && rgb.GetString() == "#123456";
            Assert.Equal(highlighted[i] == '1', colored);
        }
    }

    [Fact]
    public void CurrentPriorityNoOp_PreservesEveryFingerprint()
    {
        var sheet = CreateRules();
        var first = ReadRule(sheet, "=10");
        var other = ReadRule(sheet, "=20");
        _fixture.Send("conditionalformat.set-rule-priority", new
        {
            sheetName = sheet,
            rulePriority = first.priority,
            expectedFingerprint = first.fingerprint,
            newPriority = first.priority
        });
        Assert.Equal(first, ReadRule(sheet, "=10"));
        Assert.Equal(other, ReadRule(sheet, "=20"));
    }

    private string CreateRules()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_conditional.AddRule(_fixture.BatchToken, sheet, "A1:A5", "cellValue", "greater",
            "10", null, interiorColor: "#FFFF00").Success);
        Assert.True(_conditional.AddRule(_fixture.BatchToken, sheet, "C1:C5", "cellValue", "greater",
            "20", null, interiorColor: "#00FF00").Success);
        return sheet;
    }

    private (int priority, string fingerprint) ReadRule(string sheet, string formula)
    {
        var listed = _fixture.Send("conditionalformat.list-worksheet-rules", new { sheetName = sheet });
        using var result = JsonDocument.Parse(listed.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var rule = Assert.Single(result.RootElement.GetProperty("rules").EnumerateArray(),
            item => item.GetProperty("formula1").GetString() == formula);
        Assert.False(string.IsNullOrWhiteSpace(rule.GetProperty("fingerprint").GetString()));
        return (rule.GetProperty("priority").GetInt32(), rule.GetProperty("fingerprint").GetString()!);
    }
}
