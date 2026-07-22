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
        // Priority values, when present, should be ascending in collection order.
        var priorities = result.Rules
            .Where(r => r.Priority.HasValue)
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
}
