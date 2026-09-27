using System.Text.Json;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacPowerQueryHelperCapabilitiesTests
{
    [Fact]
    public void ParseSelectsOnlyAdvertisedPowerQueryActions()
    {
        using var document = JsonDocument.Parse(
            """
            {
              "helperVersion": "1.0.0",
              "protocolVersion": 1,
              "supportedActions": [
                "powerquery.rename",
                "powerquery.delete",
                "vba.view"
              ],
              "provenMethods": {
                "powerQueryList": false,
                "powerQueryRename": true,
                "powerQueryDelete": true
              }
            }
            """);

        var actions = MacPowerQueryHelperCapabilities.Parse(document.RootElement);

        Assert.Equal(
            ["powerquery.delete", "powerquery.rename"],
            actions.Order(StringComparer.Ordinal));
    }

    [Fact]
    public void ParseDoesNotEnableAdvertisedButUnprovenActions()
    {
        using var document = JsonDocument.Parse(
            """
            {
              "supportedActions": [
                "powerquery.list",
                "powerquery.create",
                "powerquery.update",
                "powerquery.rename",
                "powerquery.delete"
              ],
              "provenMethods": {
                "powerQueryList": false,
                "powerQueryCreate": false,
                "powerQueryUpdate": false,
                "powerQueryRename": false,
                "powerQueryDelete": false
              }
            }
            """);

        Assert.Empty(MacPowerQueryHelperCapabilities.Parse(document.RootElement));
    }

    [Fact]
    public void ParseEnablesOnlyExactExplicitCandidateAction()
    {
        using var document = JsonDocument.Parse(
            """
            {
              "supportedActions": ["powerquery.create", "powerquery.delete"],
              "provenMethods": {
                "powerQueryCreate": false,
                "powerQueryDelete": false
              }
            }
            """);

        var actions = MacPowerQueryHelperCapabilities.Parse(
            document.RootElement,
            new HashSet<string>(["powerquery.create"], StringComparer.Ordinal));

        Assert.Equal(["powerquery.create"], actions);
    }

    [Theory]
    [InlineData(null, 0)]
    [InlineData("", 0)]
    [InlineData("powerquery.create", 1)]
    [InlineData(" powerquery.create, powerquery.evaluate ", 2)]
    public void ExplicitOptInParsesExactActionNames(string? value, int expectedCount)
    {
        var actions = MacPowerQueryHelperCapabilities.ParseExplicitOptIn(value);

        Assert.Equal(expectedCount, actions.Count);
    }

    [Fact]
    public void ExplicitOptInRejectsUnknownAction()
    {
        Assert.Throws<InvalidOperationException>(() =>
            MacPowerQueryHelperCapabilities.ParseExplicitOptIn("powerquery.unknown"));
    }

    [Theory]
    [InlineData("""{"helperVersion":"1.0.0","protocolVersion":1}""")]
    [InlineData("""{"supportedActions":"powerquery.rename","provenMethods":{"powerQueryRename":true}}""")]
    [InlineData("""{"supportedActions":["powerquery.rename",1],"provenMethods":{"powerQueryRename":true}}""")]
    [InlineData("""{"supportedActions":["powerquery.rename"],"provenMethods":{"powerQueryList":false}}""")]
    [InlineData("""{"supportedActions":["powerquery.rename"],"provenMethods":{"powerQueryRename":"yes"}}""")]
    public void ParseRejectsInvalidSupportedActions(string json)
    {
        using var document = JsonDocument.Parse(json);

        Assert.Throws<InvalidOperationException>(
            () => MacPowerQueryHelperCapabilities.Parse(document.RootElement));
    }
}
