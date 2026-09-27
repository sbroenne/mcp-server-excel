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
              ]
            }
            """);

        var actions = MacPowerQueryHelperCapabilities.Parse(document.RootElement);

        Assert.Equal(
            ["powerquery.delete", "powerquery.rename"],
            actions.Order(StringComparer.Ordinal));
    }

    [Theory]
    [InlineData("""{"helperVersion":"1.0.0","protocolVersion":1}""")]
    [InlineData("""{"helperVersion":"1.0.0","protocolVersion":1,"supportedActions":"powerquery.rename"}""")]
    [InlineData("""{"helperVersion":"1.0.0","protocolVersion":1,"supportedActions":["powerquery.rename",1]}""")]
    public void ParseRejectsInvalidSupportedActions(string json)
    {
        using var document = JsonDocument.Parse(json);

        Assert.Throws<InvalidOperationException>(
            () => MacPowerQueryHelperCapabilities.Parse(document.RootElement));
    }
}
