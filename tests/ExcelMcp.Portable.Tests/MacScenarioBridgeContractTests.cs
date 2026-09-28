using System.Reflection;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacScenarioBridgeContractTests
{
    private const string ResourceName = "Sbroenne.ExcelMcp.Service.Mac.MacExcelBridge.js";

    [Theory]
    [InlineData("analysis.list-scenarios")]
    [InlineData("analysis.update-scenario")]
    [InlineData("analysis.delete-scenario")]
    [InlineData("analysis.create-scenario-summary")]
    public void DictionaryBackedScenarioActions_HaveNativeHandlers(string command)
    {
        var script = ReadBridgeScript();

        Assert.Contains($"command === \"{command}\"", script, StringComparison.Ordinal);
    }

    [Fact]
    public void ScenarioMutation_ValidatesChangingCellCountAndValues()
    {
        var script = ReadBridgeScript();

        Assert.Contains("validateScenarioValues(changingRange, args.values)", script, StringComparison.Ordinal);
        Assert.Contains("Number(changingRange.countLarge())", script, StringComparison.Ordinal);
        Assert.Contains("changingCellCount > 32", script, StringComparison.Ordinal);
        Assert.Contains("changingCellCount !== values.length", script, StringComparison.Ordinal);
    }

    [Fact]
    public void ScenarioCreateAndShow_RemainOutsideNativeBridge()
    {
        var script = ReadBridgeScript();

        Assert.DoesNotContain("command === \"analysis.create-scenario\"", script, StringComparison.Ordinal);
        Assert.DoesNotContain("command === \"analysis.show-scenario\"", script, StringComparison.Ordinal);
    }

    private static string ReadBridgeScript()
    {
        using var stream = typeof(MacAutomationHost).Assembly.GetManifestResourceStream(ResourceName);
        Assert.NotNull(stream);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
