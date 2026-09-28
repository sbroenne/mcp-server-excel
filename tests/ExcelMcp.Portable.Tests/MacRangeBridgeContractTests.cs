using System.Reflection;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacRangeBridgeContractTests
{
    private const string ResourceName = "Sbroenne.ExcelMcp.Service.Mac.MacExcelBridge.js";

    [Fact]
    public void AutoFit_UsesDeclaredApplicationCommandWithRowAndColumnRanges()
    {
        var script = ReadBridgeScript();

        Assert.Contains("excel.autofit(column)", script, StringComparison.Ordinal);
        Assert.Contains("excel.autofit(row)", script, StringComparison.Ordinal);
        Assert.Contains(
            "sheet.ranges.byName(address).columnWidth = fittedWidth",
            script,
            StringComparison.Ordinal);
        Assert.Contains(
            "sheet.ranges.byName(address).rowHeight = fittedHeight",
            script,
            StringComparison.Ordinal);
        Assert.DoesNotContain("range.columns.autofit()", script, StringComparison.Ordinal);
        Assert.DoesNotContain("range.rows.autofit()", script, StringComparison.Ordinal);
    }

    [Fact]
    public void MergeMutations_UseDeclaredCommandsAndVerifyFreshRangeState()
    {
        var script = ReadBridgeScript();

        Assert.Contains("excel.merge(range)", script, StringComparison.Ordinal);
        Assert.Contains("excel.unmerge(range)", script, StringComparison.Ordinal);
        Assert.Contains(
            "sheet.ranges.byName(args.rangeAddress).mergeCells()",
            script,
            StringComparison.Ordinal);
        Assert.DoesNotContain("range.merge()", script, StringComparison.Ordinal);
        Assert.DoesNotContain("range.unmerge()", script, StringComparison.Ordinal);
    }

    private static string ReadBridgeScript()
    {
        using var stream = typeof(MacAutomationHost).Assembly.GetManifestResourceStream(ResourceName);
        Assert.NotNull(stream);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
