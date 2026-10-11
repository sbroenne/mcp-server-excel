using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "Drawing")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativeDrawingAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task NativeGroupAlignmentDistributionAndOrder_KeepActualObjects()
    {
        foreach (var (name, left, top) in new[] { ("DrawFirst", "20", "20"), ("DrawSecond", "100", "70") })
        {
            var created = await CommandAsync("drawing", "add-shape", "--sheet-name", "Data", "--name", name,
                "--left", left, "--top", top, "--width", "40", "--height", "30");
            Assert.Equal(name, Text(created, "drawingObject.name"));
        }
        var duplicate = await CommandAsync("drawing", "duplicate-object", "--sheet-name", "Data", "--object-name", "DrawFirst",
            "--new-name", "DrawThird", "--offset-left", "180", "--offset-top", "70");
        Assert.Equal("DrawThird", Text(duplicate, "drawingObjects.0.name"));
        Assert.Equal(200, Number(duplicate, "drawingObjects.0.left"));
        var grouped = await CommandAsync("drawing", "group-objects", "--sheet-name", "Data",
            "--object-names", """["DrawFirst","DrawSecond"]""", "--group-name", "DrawGroup");
        Assert.Equal("DrawGroup", Text(grouped, "drawingObjects.0.name"));
        Assert.Equal(2, Items(grouped, "drawingObjects.0.children").Length);
        var read = await CommandAsync("drawing", "get-object", "--sheet-name", "Data", "--object-name", "DrawGroup");
        Assert.Equal(2, Items(read, "drawingObject.children").Length);
        var ungrouped = Items(await CommandAsync("drawing", "ungroup-object", "--sheet-name", "Data", "--object-name", "DrawGroup"), "drawingObjects");
        Assert.Equal(2, ungrouped.Length);
        Assert.Contains(ungrouped, item => Text(item, "name") == "DrawFirst");
        const string names = """["DrawFirst","DrawSecond","DrawThird"]""";
        var aligned = Items(await CommandAsync("drawing", "align-objects", "--sheet-name", "Data", "--object-names", names,
            "--alignment", "Top"), "drawingObjects");
        Assert.Equal(3, aligned.Length);
        Assert.Single(aligned.Select(item => Number(item, "top")).Distinct());
        var distributed = await CommandAsync("drawing", "distribute-objects", "--sheet-name", "Data", "--object-names", names,
            "--distribution", "Horizontal");
        Assert.Equal(110, Number(distributed, "drawingObjects.1.left"));
        var ordered = await CommandAsync("drawing", "set-z-order", "--sheet-name", "Data", "--object-name", "DrawThird", "--z-order", "SendToBack");
        Assert.Equal(1, Number(ordered, "drawingObjects.0.zOrderPosition"));
    }
}
