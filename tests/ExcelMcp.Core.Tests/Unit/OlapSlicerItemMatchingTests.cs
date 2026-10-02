using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "Slicer")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class OlapSlicerItemMatchingTests
{
    [Fact]
    public void AmbiguousCaption_ErrorProvidesNamesThatCanBeSelected()
    {
        PivotTableCommands.SlicerItemState[] items =
        [
            new("[Calendar].[Quarter].&[2025]&[Q1]", "Q1", true),
            new("[Calendar].[Quarter].&[2026]&[Q1]", "q1", false),
            new("[Calendar].[Quarter].&[2026]&[Q2]", "Q2", false)
        ];

        var error = Assert.Throws<ArgumentException>(() =>
            PivotTableCommands.ResolveOlapSlicerItemName(items, "q1"));

        Assert.Equal("requested", error.ParamName);
        Assert.Contains("ambiguous", error.Message);
        var matchingNames = items.Take(2).Select(item => item.Name).ToArray();
        string candidatesJson = JsonSerializer.Serialize(matchingNames);
        Assert.Contains(candidatesJson, error.Message);
        const string marker = "matching MDX unique names: ";
        string returnedJson = error.Message
            [(error.Message.IndexOf(marker, StringComparison.Ordinal) + marker.Length)..(error.Message.LastIndexOf(']') + 1)];
        var candidates = JsonSerializer.Deserialize<string[]>(returnedJson)!;
        Assert.Equal(matchingNames, candidates);
        Assert.All(candidates, name =>
            Assert.Equal(name, PivotTableCommands.ResolveOlapSlicerItemName(items, name)));
    }

    [Theory]
    [InlineData("q2", "[Calendar].[Quarter].&[Q2]")]
    [InlineData("[calendar].[quarter].&[q1]", "[Calendar].[Quarter].&[Q1]")]
    public void CaptionOrUniqueName_ResolvesCaseInsensitively(string requested, string expected)
    {
        PivotTableCommands.SlicerItemState[] items =
        [
            new("[Calendar].[Quarter].&[Q1]", "Q1", true),
            new("[Calendar].[Quarter].&[Q2]", "Q2", false)
        ];
        Assert.Equal(expected, PivotTableCommands.ResolveOlapSlicerItemName(items, requested));
    }

    [Fact]
    public void UniqueName_TakesPrecedenceOverAnotherItemsCaption()
    {
        PivotTableCommands.SlicerItemState[] items =
        [
            new("member-1", "Q1", true),
            new("member-2", "member-1", false)
        ];
        Assert.Equal("member-1", PivotTableCommands.ResolveOlapSlicerItemName(items, "member-1"));
    }

    [Fact]
    public void UnknownCaption_RemainsExplicitError()
    {
        PivotTableCommands.SlicerItemState[] items = [new("member-1", "Q1", true)];
        var error = Assert.Throws<ArgumentException>(() =>
            PivotTableCommands.ResolveOlapSlicerItemName(items, "missing"));
        Assert.Equal("requested", error.ParamName);
        Assert.Contains("was not found", error.Message);
    }
}
