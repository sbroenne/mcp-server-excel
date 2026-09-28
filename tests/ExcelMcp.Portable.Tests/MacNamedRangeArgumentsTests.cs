using System.Globalization;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacNamedRangeArgumentsTests
{
    [Theory]
    [InlineData("list")]
    [InlineData("create")]
    [InlineData("read")]
    [InlineData("write")]
    [InlineData("update")]
    [InlineData("delete")]
    public void Acceptance_RequiresExactOptInAndProductionRemainsGated(string action)
    {
        var command = $"namedrange.{action}";
        var capability = MacCommandCapabilities.Get(command);
        Assert.False(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.True(MacNamedRangeArguments.CanUseForAcceptance(command, "1"));
        Assert.False(MacNamedRangeArguments.CanUseForAcceptance(command, null));
        Assert.False(MacNamedRangeArguments.CanUseForAcceptance(command, "true"));
        Assert.False(MacNamedRangeArguments.CanUseForAcceptance(command, " 1"));
    }

    [Theory]
    [InlineData("namedrange.unknown")]
    [InlineData("namedrange.CREATE")]
    [InlineData("range.get-current-region")]
    [InlineData("vba.run")]
    public void Acceptance_DoesNotBroadenToUnknownOrOtherCommands(string command) =>
        Assert.False(MacNamedRangeArguments.CanUseForAcceptance(command, "1"));

    [Theory]
    [InlineData("create", "Data!$A$1")]
    [InlineData("update", "==Data!$A$1")]
    public void Prepare_NormalizesReferenceLikeWindows(string action, string reference)
    {
        var args = new JsonObject { ["name"] = "Revenue", ["reference"] = reference };
        MacNamedRangeArguments.Prepare(action, args);
        Assert.Equal("=Data!$A$1", args["reference"]!.GetValue<string>());
    }

    [Theory]
    [InlineData("create", "")]
    [InlineData("update", "  ")]
    public void Prepare_RejectsInvalidNameBeforeAutomation(string action, string name)
    {
        var args = new JsonObject { ["name"] = name, ["reference"] = "Data!A1" };
        Assert.Throws<ArgumentException>(() => MacNamedRangeArguments.Prepare(action, args));
        args["name"] = new string('x', 256);
        Assert.Throws<ArgumentException>(() => MacNamedRangeArguments.Prepare(action, args));
    }

    [Fact]
    public void Prepare_ConvertsNumericBooleanAndTextWritesUsingCallerCulture()
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            var args = new JsonObject { ["name"] = "Input", ["value"] = "12,5" };
            MacNamedRangeArguments.Prepare("write", args);
            Assert.Equal(12.5, args["parsedValue"]!.GetValue<double>());
            args["value"] = "TRUE";
            MacNamedRangeArguments.Prepare("write", args);
            Assert.True(args["parsedValue"]!.GetValue<bool>());
            args["value"] = "";
            MacNamedRangeArguments.Prepare("write", args);
            Assert.Equal("", args["parsedValue"]!.GetValue<string>());
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }

    [Fact]
    public void Prepare_RejectsMissingOrIncorrectlyTypedArguments()
    {
        Assert.Throws<ArgumentException>(() =>
            MacNamedRangeArguments.Prepare("read", new JsonObject()));
        Assert.Throws<ArgumentException>(() =>
            MacNamedRangeArguments.Prepare("write", new JsonObject { ["name"] = "Input", ["value"] = 7 }));
        Assert.Throws<ArgumentException>(() =>
            MacNamedRangeArguments.Prepare("unknown", new JsonObject()));
    }
}
