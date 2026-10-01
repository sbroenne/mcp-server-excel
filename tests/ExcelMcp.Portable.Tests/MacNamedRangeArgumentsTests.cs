using System.Globalization;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service.Mac;
using Sbroenne.ExcelMcp.Generated;
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
    public void ProvenActions_AreProductionEnabledWithoutAcceptanceOverride(string action)
    {
        var command = $"namedrange.{action}";
        var capability = MacCommandCapabilities.Get(command);
        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Equal("Implemented", capability.ImplementationStatus);
        Assert.Contains("Real CLI/MCP acceptance", capability.Evidence, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("namedrange.unknown")]
    [InlineData("namedrange.CREATE")]
    [InlineData("namedrange.Read")]
    [InlineData("namedrange.create-extra")]
    public void UnknownActions_RemainUnavailableAndFailPreparation(string command)
    {
        Assert.False(MacCommandCapabilities.Get(command).IsAvailable);
        Assert.Throws<ArgumentException>(() =>
            MacNamedRangeArguments.Prepare(command["namedrange.".Length..], new JsonObject()));
    }

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
    public void Prepare_UsesInvariantNumericBooleanAndTextWrites()
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            var args = new JsonObject { ["name"] = "Input", ["value"] = "12,5" };
            MacNamedRangeArguments.Prepare("write", args);
            Assert.Equal("12,5", args["parsedValue"]!.GetValue<string>());
            args["value"] = "12.5";
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

    [Theory]
    [InlineData("get-values")]
    [InlineData("set-values")]
    public void BulkRangeBinding_RequiresCorrespondingCapabilityAndUsesRequestedName(string action)
    {
        var args = new JsonObject { ["sheetName"] = "", ["rangeAddress"] = "Input" };
        Assert.Throws<PlatformNotSupportedException>(() =>
            MacNamedRangeArguments.PrepareRangeBinding(action, args, enabled: false));
        MacNamedRangeArguments.PrepareRangeBinding(action, args, enabled: true);
        Assert.Equal("Input", args["namedRangeName"]!.GetValue<string>());
    }

    [Fact]
    public void BulkRangeBinding_DoesNotTrustCallerSuppliedInternalName()
    {
        var args = new JsonObject
        {
            ["sheetName"] = "Data",
            ["rangeAddress"] = "A1",
            ["namedRangeName"] = "UnrequestedName"
        };
        MacNamedRangeArguments.PrepareRangeBinding("set-values", args, enabled: false);
        Assert.False(args.ContainsKey("namedRangeName"));
        args["sheetName"] = "";
        args["rangeAddress"] = " ";
        Assert.Throws<ArgumentException>(() =>
            MacNamedRangeArguments.PrepareRangeBinding("set-values", args, enabled: true));
    }

    [Theory]
    [InlineData("get-values")]
    [InlineData("set-values")]
    public void SharedRangeContract_AllowsExplicitEmptySheetButStillRequiresAString(string action)
    {
        var args = new JsonObject { ["sheetName"] = "", ["rangeAddress"] = "Input" };
        if (action == "set-values") args["values"] = JsonNode.Parse("[[17]]");
        ServiceRegistry.Range.ValidateActionArguments(action, args.ToJsonString());
        args.Remove("sheetName");
        Assert.Throws<ArgumentException>(() => ServiceRegistry.Range.ValidateActionArguments(action, args.ToJsonString()));
        args["sheetName"] = null;
        Assert.Throws<ArgumentException>(() => ServiceRegistry.Range.ValidateActionArguments(action, args.ToJsonString()));
        args["sheetName"] = 1;
        Assert.Throws<ArgumentException>(() => ServiceRegistry.Range.ValidateActionArguments(action, args.ToJsonString()));
    }
}
