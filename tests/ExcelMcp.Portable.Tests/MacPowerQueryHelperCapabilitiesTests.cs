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
                "powerQueryMutation": true
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
                "powerQueryMutation": false
              }
            }
            """);

        Assert.Empty(MacPowerQueryHelperCapabilities.Parse(document.RootElement));
    }

    [Theory]
    [InlineData("""{"helperVersion":"1.0.0","protocolVersion":1}""")]
    [InlineData("""{"supportedActions":"powerquery.rename","provenMethods":{"powerQueryList":false,"powerQueryMutation":true}}""")]
    [InlineData("""{"supportedActions":["powerquery.rename",1],"provenMethods":{"powerQueryList":false,"powerQueryMutation":true}}""")]
    [InlineData("""{"supportedActions":["powerquery.rename"],"provenMethods":{"powerQueryList":false}}""")]
    [InlineData("""{"supportedActions":["powerquery.rename"],"provenMethods":{"powerQueryList":false,"powerQueryMutation":"yes"}}""")]
    public void ParseRejectsInvalidSupportedActions(string json)
    {
        using var document = JsonDocument.Parse(json);

        Assert.Throws<InvalidOperationException>(
            () => MacPowerQueryHelperCapabilities.Parse(document.RootElement));
    }
}
