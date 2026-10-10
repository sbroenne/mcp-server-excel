using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "ConditionalFormat")]
[Trait("Feature", "FineFormatting")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativeConditionalFormatAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task SelectedRuleAndFormattingOnlyPaste_PreserveStateAfterReopen()
    {
        await CommandAsync("rangeformat", "format", "--sheet-name", "Data", "--range-addresses", "A1:A2",
            "--range-addresses", "C1:C2", "--format-options", """{"bold":true,"fillColor":"#FFFF00"}""");
        var format = await RangeAsync("rangeformat", "get-format", "A1:C2", "--view", "both");
        Assert.Equal(6, Number(format, "cellCount"));
        Assert.Equal(6, Items(format, "cells").Length);
        Assert.True(Bool(format, "cells.0.stored.font.bold"));
        Assert.Equal("#FFFF00", Text(format, "cells.0.stored.fill.color.rgb"));
        Assert.Equal("#FFFF00", Text(format, "cells.0.displayed.fill.color.rgb"));
        await RangeAsync("conditionalformat", "add-rule", "B1:B10", "--rule-type", "top10", "--rank", "7",
            "--top10-percent", "true", "--font-bold", "true", "--font-italic", "false");
        var rule = Assert.Single(Items(await RangeAsync("conditionalformat", "list-rules", "B1:B10"), "rules"));
        Assert.Equal(7, Number(rule, "top10.rank"));
        Assert.True(Bool(rule, "top10.percent"));
        var updated = await SelectedAsync("update-rule", rule, "--options", """{"rank":5,"stopIfTrue":false,"appliesTo":"B1:B8"}""");
        var selected = Assert.Single(Items(updated, "rules"));
        Assert.Equal(5, Number(selected, "top10.rank"));
        Assert.False(Bool(selected, "stopIfTrue"));
        Assert.Equal("$B$1:$B$8", Text(selected, "appliesTo"));
        await RangeAsync("conditionalformat", "add-rule", "W10:W12", "--rule-type", "expression",
            "--formula1", "=TRUE()", "--stop-if-true", "false");
        var before = await CommandAsync("conditionalformat", "list-worksheet-rules", "--sheet-name", "Data");
        Assert.Equal(2, Items(before, "rules").Length);
        var disposable = Assert.Single(Items(before, "rules"), item => Text(item, "type") == "expression");
        var moved = await SelectedAsync("set-rule-priority", disposable, "--new-priority", "1");
        Assert.Equal(2, Items(moved, "rules").Length);
        disposable = Assert.Single(Items(moved, "rules"), item => Text(item, "type") == "expression");
        Assert.Equal(1, Number(disposable, "priority"));
        var afterDelete = await SelectedAsync("delete-rule", disposable);
        Assert.Equal(5, Number(Assert.Single(Items(afterDelete, "rules")), "top10.rank"));
        await ValuesAsync("H1", "[[7]]");
        var copied = await CommandAsync("range", "copy", "--source-sheet", "Data", "--source-range", "A1:C2",
            "--target-sheet", "Data", "--target-range", "H1", "--paste-kind", "formats");
        Assert.Equal("$H$1:$J$2", Text(copied, "destinationAddress"));
        Assert.Equal("formats", Text(copied, "pasteKind"));
        Assert.Equal(7, Number(await RangeAsync("range", "get-values", "H1"), "values.0.0"));
        var pasted = await RangeAsync("rangeformat", "get-format", "H1");
        Assert.True(Bool(pasted, "cells.0.stored.font.bold"));
        Assert.Equal("#FFFF00", Text(pasted, "cells.0.stored.fill.color.rgb"));
        await ReopenAsync();
        var saved = Items(await CommandAsync("conditionalformat", "list-worksheet-rules", "--sheet-name", "Data"), "rules");
        Assert.Equal(2, saved.Length);
        foreach (var address in new[] { "$B$1:$B$8", "$I$1:$I$2" })
        {
            var persisted = Assert.Single(saved, item => Text(item, "appliesTo") == address);
            Assert.Equal(5, Number(persisted, "top10.rank"));
            Assert.False(Bool(persisted, "stopIfTrue"));
        }
    }

    private Task<System.Text.Json.JsonElement> SelectedAsync(string action, System.Text.Json.JsonElement rule, params string[] arguments) =>
        CommandAsync("conditionalformat", action, ["--sheet-name", "Data", "--rule-priority",
            Number(rule, "priority").ToString(System.Globalization.CultureInfo.InvariantCulture),
            "--expected-fingerprint", Text(rule, "fingerprint"), .. arguments]);
}
